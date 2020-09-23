VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSchoolInfo 
   Caption         =   "School Info & Current Enrollment"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmSchoolInfo.frx":0000
   ScaleHeight     =   5925
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin MSDataListLib.DataCombo dcSemester 
      Height          =   315
      Left            =   2115
      TabIndex        =   26
      Top             =   4950
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcSy 
      Height          =   315
      Left            =   2115
      TabIndex        =   25
      Top             =   4500
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin VB.TextBox txtFaxNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2115
      MaxLength       =   15
      TabIndex        =   16
      Top             =   3555
      Width           =   2130
   End
   Begin VB.TextBox txtTelNo2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2115
      MaxLength       =   15
      TabIndex        =   15
      Top             =   3150
      Width           =   2130
   End
   Begin VB.TextBox txtTelNo1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2115
      MaxLength       =   15
      TabIndex        =   14
      Top             =   2745
      Width           =   2130
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2115
      MaxLength       =   60
      TabIndex        =   2
      Top             =   2340
      Width           =   5775
   End
   Begin VB.TextBox txtSchoolName2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2115
      MaxLength       =   35
      TabIndex        =   1
      Top             =   1935
      Width           =   4470
   End
   Begin VB.TextBox txtSchoolName1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2115
      MaxLength       =   30
      TabIndex        =   0
      Top             =   1530
      Width           =   4470
   End
   Begin VB.PictureBox picFooter 
      Align           =   2  'Align Bottom
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   8640
      TabIndex        =   8
      Top             =   5430
      Width           =   8700
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "Close"
         Height          =   315
         Left            =   7485
         TabIndex        =   5
         Top             =   60
         Width           =   1035
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H000080FF&
         Caption         =   "Delete"
         Height          =   315
         Left            =   6405
         MaskColor       =   &H000080FF&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   60
         Width           =   1035
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   315
         Left            =   5325
         TabIndex        =   3
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
         TabIndex        =   9
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
      ScaleWidth      =   8700
      TabIndex        =   6
      Top             =   0
      Width           =   8700
      Begin VB.Image Image2 
         Height          =   375
         Left            =   8325
         Picture         =   "frmSchoolInfo.frx":10380
         Top             =   -45
         Width           =   390
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "School Information"
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
         Left            =   1020
         TabIndex        =   7
         Top             =   120
         Width           =   4845
      End
      Begin VB.Image Image1 
         Height          =   780
         Left            =   0
         Picture         =   "frmSchoolInfo.frx":10830
         Top             =   60
         Width           =   915
      End
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "location"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   2565
      TabIndex        =   24
      Top             =   1080
      Width           =   1995
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "enrollment"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   2025
      TabIndex        =   23
      Top             =   4005
      Width           =   1995
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Registered"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   465
      Left            =   540
      TabIndex        =   22
      Top             =   1080
      Width           =   2040
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Current"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   465
      Left            =   540
      TabIndex        =   21
      Top             =   4005
      Width           =   1500
   End
   Begin VB.Shape shpBar 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   0
      Top             =   0
      Width           =   6915
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Accounts"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   210
      Left            =   75
      TabIndex        =   20
      Top             =   0
      Width           =   4815
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Semester"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   19
      Top             =   4995
      Width           =   1230
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "School Year"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   720
      TabIndex        =   18
      Top             =   4500
      Width           =   1230
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax No"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   765
      TabIndex        =   17
      Top             =   3555
      Width           =   1230
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   765
      TabIndex        =   13
      Top             =   2385
      Width           =   1230
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone No"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   765
      TabIndex        =   12
      Top             =   2790
      Width           =   1230
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User Id:"
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
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   1230
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "School Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   765
      TabIndex        =   10
      Top             =   1575
      Width           =   1230
   End
End
Attribute VB_Name = "frmSchoolInfo"
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

Private Sub cmdDelete_Click()
    If cmdSave.Caption = "Update" Then
        If MsgBox("Do you want to Delete the record?", vbYesNo + vbCritical, Me.Caption) = vbYes Then
            RunSql ("TRUNCATE TABLE dbo.SchoolInfo;")
            rsData.Requery
            ClearEntries Me
            txtSchoolName1.SetFocus
        End If
    Else
        MsgBox ("No record selected.")
    End If
End Sub

Private Sub cmdSave_Click()
    If EntriesValid Then
        If MsgBox("Do you want to " & cmdSave.Caption & " record?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
            strSql = "INSERT INTO dbo.SchoolInfo(SchoolName1, SchoolName2, Address, TelNo1, TelNo2, Faxno, CurrentSyId, CurrentSemesterId) " & _
                                    "VALUES('" & txtSchoolName1.Text & "', " & _
                                           "'" & txtSchoolName1.Text & "', " & _
                                           "'" & txtAddress.Text & "', " & _
                                           "'" & txtTelNo1.Text & "', " & _
                                           "'" & txtTelNo2.Text & "', " & _
                                           "'" & txtFaxNo.Text & "', " & _
                                           "'" & dcSy.BoundText & "', " & _
                                           "'" & dcSemester.BoundText & "');"
            If cmdSave.Caption = "Update" Then
                strSql = "UPDATE dbo.SchoolInfo SET SchoolName1 = '" & txtSchoolName1.Text & "', " & _
                                              "SchoolName2 = '" & txtSchoolName2.Text & "', " & _
                                              "Address = '" & txtAddress.Text & "', " & _
                                              "TelNo1 = '" & txtTelNo1.Text & "', " & _
                                              "TelNo2 = '" & txtTelNo2.Text & "', " & _
                                              "FaxNo = '" & txtFaxNo.Text & "', " & _
                                              "CurrentSyId = '" & dcSy.BoundText & "', " & _
                                              "CurrentSemesterId = '" & dcSemester.BoundText & "';"
            End If
            RunSql (strSql)
            If rsData.State = adStateOpen Then
                rsData.Requery
            End If
            'ClearEntries Me
            MsgBox "School Info & Current Enrollment was updated successfully.", vbInformation, Me.Caption
            txtSchoolName1.SetFocus
        End If
    End If
End Sub


Private Sub Form_Activate()
    MakeTransparent Me.hwnd, 255 'Unfade form
End Sub

Private Sub Form_Deactivate()
    MakeTransparent Me.hwnd, 190 'Fade Form
End Sub

Private Sub Form_Load()
    'Bind School Year
    BindDataCombo "SELECT * FROM dbo.Sy", "Sy", dcSy, "SyId", True
    'Bind Semester
    BindDataCombo "SELECT * FROM dbo.Semester", "Semester", dcSemester, "SemesterId", True
    
    'On Error Resume Next
    ValidateAccessLevel Me, "Save"
    strSql = "SELECT * FROM SchoolInfo;"
    
    Set rsData = GetRecordset(strSql)
    If rsData.State = adStateOpen Then
        If rsData.RecordCount > 0 Then
            txtSchoolName1.Text = rsData.Fields("SchoolName1").Value
            txtSchoolName2.Text = rsData.Fields("SchoolName2").Value
            txtAddress.Text = rsData.Fields("Address").Value
            txtTelNo1.Text = rsData.Fields("TelNo1").Value
            txtTelNo2.Text = rsData.Fields("TelNo2").Value
            txtFaxNo.Text = rsData.Fields("FaxNo").Value
            dcSy.BoundText = rsData.Fields("CurrentSyId").Value
            dcSemester.BoundText = rsData.Fields("CurrentSemesterId").Value
            ValidateAccessLevel Me, "Update"
        End If
    End If
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsData.Close
End Sub

Private Function EntriesValid() As Boolean
    EntriesValid = False
    If txtSchoolName1.Text = "" Then
        MsgBox ("School Name required!")
        txtSchoolName1.SetFocus
        Exit Function
    End If
    If txtSchoolName2.Text = "" Then
        MsgBox ("Sub-School Name required!")
        txtSchoolName2.SetFocus
        Exit Function
    End If
    If txtAddress.Text = "" Then
        MsgBox ("School Address required!")
        txtAddress.SetFocus
        Exit Function
    End If
    If txtTelNo1.Text = "" Then
        MsgBox ("Telephone Number required!")
        txtTelNo1.SetFocus
        Exit Function
    End If
'    If cboSy.Text = "" Then
'        MsgBox ("Course Year required!")
'        cboCourseYear.SetFocus
'        Exit Function
'    End If
    EntriesValid = True
End Function


'Private Sub txtCourseCode_LostFocus()
'    If RecordExist("CourseCode", "dbo.SchoolInfo", txtCourseCode.Text) Then
'        If MsgBox("Course Code: " & txtCourseCode.Text & " already exist. Would you like to retrieve the information?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
'            'call GetRecordset Function
'            Dim rs As New ADODB.Recordset
'            Set rs = GetRecordset("SELECT * FROM SchoolInfo WHERE coursecode = '" & txtCourseCode.Text & "'")
'
'            If rs.State = adStateOpen Then
'                If Not rs.BOF Then 'if userid exists
'                    txtCourseDesc.Text = rs!CourseDesc
'                    txtCollege.Text = rs!College
'                    cboCourseYear.Text = rs!CourseYear
'                    ValidateAccessLevel Me, "Update"
'                End If
'            End If
'        Else
'            ClearEntries Me
'            txtCourseCode.SetFocus
'        End If
'    End If
'End Sub

Private Sub txtSchoolName1_GotFocus()
    FocusMe (txtSchoolName1)
End Sub

Private Sub txtSchoolName2_GotFocus()
    FocusMe (txtSchoolName1)
End Sub

Private Sub txtAddress_GotFocus()
    FocusMe (txtAddress)
End Sub

Private Sub txtTelNo1_GotFocus()
    FocusMe (txtTelNo1)
End Sub

Private Sub txtTelNo2_GotFocus()
    FocusMe (txtTelNo2)
End Sub

Private Sub txtFaxNo_GotFocus()
    FocusMe (txtFaxNo)
End Sub



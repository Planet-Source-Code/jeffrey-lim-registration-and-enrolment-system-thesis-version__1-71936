VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DCLC-RES - Login"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4395
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picFooter 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   4365
      TabIndex        =   9
      Top             =   2070
      Width           =   4395
      Begin VB.CommandButton cmdOk 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   350
         Left            =   3360
         TabIndex        =   11
         Top             =   60
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   350
         Left            =   2340
         TabIndex        =   10
         Top             =   60
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00100F0D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   4395
      TabIndex        =   7
      Top             =   0
      Width           =   4395
      Begin VB.Image Image1 
         Height          =   645
         Left            =   60
         Picture         =   "frmLogin.frx":22952
         Top             =   60
         Width           =   750
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Plese enter your username and password  to login."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   900
         TabIndex        =   8
         Top             =   135
         Width           =   3135
      End
   End
   Begin VB.TextBox txtServer 
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
      ForeColor       =   &H00000000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1260
      TabIndex        =   5
      Top             =   1620
      Width           =   2775
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1260
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1230
      Width           =   2775
   End
   Begin VB.TextBox txtUsername 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1260
      TabIndex        =   0
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "SQL &Server:"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   2
      Left            =   270
      TabIndex        =   6
      Top             =   1620
      Width           =   945
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1230
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   270
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    'declare recordset variable
    Dim rs As ADODB.Recordset

    'call GetRecordset Function
    Set rs = GetRecordset("SELECT * FROM users WHERE userid = '" & txtUsername.Text & "'")

    If rs.State = adStateOpen Then
        If Not rs.BOF Then 'if userid exists
           If rs!Password = txtPassword.Text Then 'check for correct password
                With MDIForm1.StatusBar1.Panels
                    .Item(2).Text = rs!UserName
                    .Item(5).Text = Now
                End With
                User.UserId = txtUsername.Text
                User.UserName = rs!UserName
                User.UserType = rs!UserType
                
                'get School info and current sy
                Set rs = GetRecordset("SELECT * FROM dbo.SchoolInfo AS si " & _
                                      "INNER JOIN dbo.Sy AS sy ON (si.CurrentSyId = sy.SyId) " & _
                                      "INNER JOIN dbo.Semester AS sem ON (si.CurrentSemesterId = sem.SemesterId) " & _
                                      ";")
                If rs.State = adStateOpen Then
                    If rs.RecordCount > 0 Then
                        SchoolInformation.SchoolName1 = rs!SchoolName1
                        SchoolInformation.SchoolName2 = rs!SchoolName2
                        SchoolInformation.Address = rs!Address
                        SchoolInformation.TelNo1 = rs!TelNo1
                        SchoolInformation.TelNo2 = rs!TelNo2
                        SchoolInformation.FaxNo = rs!FaxNo
                        SchoolInformation.CurrentSyId = rs!CurrentSyId
                        SchoolInformation.CurrentSemesterId = rs!CurrentSemesterId
                        SchoolInformation.Sy = rs!Sy
                        SchoolInformation.Semester = rs!Semester
                        With MDIForm1.StatusBar1.Panels
                            .Item(10).Text = rs!Sy
                            .Item(12).Text = rs!Semester
                        End With
                    End If
                End If
    
                Set rs = Nothing
                Unload Me
           Else
                MsgBox "Invalid Password, try again!", , "DCLC - Login"
                txtPassword.SetFocus
                SendKeys "{Home}+{End}"
           End If
        Else
           MsgBox "Invalid Userid, try again!", , "DCLC - Login"
           txtUsername.SetFocus
           SendKeys "{Home}+{End}"
        End If
        'rs.Close
    Else
        MsgBox ("GetRecordset error (closed)")
    End If
    
End Sub

Private Sub cmdCancel_Click()
    End
End Sub



Private Sub txtPassword_GotFocus()
    FocusMe txtPassword
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub

Private Sub txtUsername_GotFocus()
    FocusMe txtUsername
End Sub

Private Sub txtUsername_KeyPress(KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub

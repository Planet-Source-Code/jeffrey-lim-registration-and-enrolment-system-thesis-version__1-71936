VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmUsers 
   Caption         =   "System Users"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmUsers.frx":0000
   ScaleHeight     =   5145
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboUserType 
      Height          =   315
      ItemData        =   "frmUsers.frx":10380
      Left            =   1665
      List            =   "frmUsers.frx":1038D
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2430
      Width           =   2940
   End
   Begin VB.TextBox txtPassword 
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
      Left            =   1665
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2025
      Width           =   2715
   End
   Begin VB.TextBox txtUsername 
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
      Left            =   1665
      MaxLength       =   35
      TabIndex        =   1
      Top             =   1620
      Width           =   4245
   End
   Begin VB.TextBox txtUserId 
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
      Left            =   1665
      MaxLength       =   10
      TabIndex        =   0
      Top             =   1215
      Width           =   2085
   End
   Begin MSDataGridLib.DataGrid dgUsers 
      Height          =   1455
      Left            =   45
      TabIndex        =   7
      Top             =   3105
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   2566
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picFooter 
      Align           =   2  'Align Bottom
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   6600
      TabIndex        =   10
      Top             =   4650
      Width           =   6660
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "Close"
         Height          =   315
         Left            =   5460
         TabIndex        =   6
         Top             =   60
         Width           =   1035
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H000080FF&
         Caption         =   "Delete"
         Height          =   315
         Left            =   4380
         MaskColor       =   &H000080FF&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         Width           =   1035
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   315
         Left            =   3300
         TabIndex        =   4
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
         TabIndex        =   11
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
      ScaleWidth      =   6660
      TabIndex        =   8
      Top             =   0
      Width           =   6660
      Begin VB.Image Image2 
         Height          =   375
         Left            =   7560
         Picture         =   "frmUsers.frx":103B4
         Top             =   0
         Width           =   390
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Users"
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
         TabIndex        =   9
         Top             =   120
         Width           =   3135
      End
      Begin VB.Image Image1 
         Height          =   780
         Left            =   0
         Picture         =   "frmUsers.frx":10864
         Top             =   60
         Width           =   915
      End
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
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
      Left            =   315
      TabIndex        =   16
      Top             =   1665
      Width           =   1230
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   315
      TabIndex        =   15
      Top             =   2070
      Width           =   1230
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Access Level:"
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
      Left            =   315
      TabIndex        =   14
      Top             =   2475
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
      TabIndex        =   13
      Top             =   0
      Width           =   1230
   End
   Begin VB.Label Label1 
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
      Height          =   285
      Left            =   315
      TabIndex        =   12
      Top             =   1260
      Width           =   1230
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsData As New ADODB.Recordset
Dim strSyId As String

'Icon on Message vbCritical 16, vbQuestion 32, vbExclamation 48, vbInformation 64
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If cmdSave.Caption = "Update" Then
        If MsgBox("Do you want to Delete the record?", vbYesNo + vbCritical, Me.Caption) = vbYes Then
            RunSql ("DELETE FROM dbo.Users WHERE UserId = '" + txtUserId.Text & "';")
            rsData.Requery
            ClearEntries Me
            txtUserId.SetFocus
        End If
    Else
        MsgBox ("No record selected.")
    End If
End Sub

Private Sub cmdSave_Click()
    If EntriesValid Then
        If MsgBox("Do you want to " & cmdSave.Caption & " record?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
            strSql = "INSERT INTO dbo.Users(UserId, Username, Password, UserType) " & _
                                    "VALUES('" & txtUserId.Text & "', " & _
                                           "'" & txtUsername.Text & "', " & _
                                           "'" & txtPassword.Text & "', " & _
                                           "'" & cboUserType.Text & "');"
            If cmdSave.Caption = "Update" Then
                strSql = "UPDATE dbo.Users SET Username = '" & txtUsername.Text & "', " & _
                                              "Password = '" & txtPassword.Text & "', " & _
                                              "UserType = '" & cboUserType.Text & "' " & _
                                        "WHERE UserId = '" & txtUserId.Text & "';"
            End If
            RunSql (strSql)
            rsData.Requery
            ClearEntries Me
            txtUserId.SetFocus
        End If
    End If
End Sub

Private Sub dgUsers_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    txtUserId = rsData.Fields(0).Value
    txtUsername = rsData.Fields(1).Value
    txtPassword = rsData.Fields(2).Value
    cboUserType.Text = rsData.Fields(3).Value
    ValidateAccessLevel Me, "Update"
End Sub

Private Sub Form_Activate()
    MakeTransparent Me.hwnd, 255 'Unfade form
End Sub

Private Sub Form_Deactivate()
    MakeTransparent Me.hwnd, 190 'Fade Form
End Sub

Private Sub Form_Load()
    'On Error Resume Next
    strSql = "SELECT UserId, Username, Password, UserType AS Access FROM users WHERE UserStatus = 'A' ORDER BY Username;"
    '-- test access level here
    If LCase(User.UserType) <> "administrator" Then
        strSql = "SELECT UserId, Username, Password, UserType AS Access FROM users WHERE UserStatus = 'A' AND UserId = '" & User.UserId & "' ORDER BY Username;"
        cboUserType.Enabled = False
    End If
    
    Set rsData = GetRecordset(strSql)
    If rsData.State = adStateOpen Then
        Set dgUsers.DataSource = rsData
        'dgUsers.Columns.Item(0).Caption = "ID"
    End If
    ValidateAccessLevel Me, "Insert"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsData.Close
End Sub

Public Function EntriesValid() As Boolean
    EntriesValid = False
    If txtUserId.Text = "" Then
        MsgBox ("User Id required!")
        txtUserId.SetFocus
        Exit Function
    End If
    If txtUsername.Text = "" Then
        MsgBox ("Username required!")
        txtUsername.SetFocus
        Exit Function
    End If
    If txtPassword.Text = "" Then
        MsgBox ("Password required!")
        txtPassword.SetFocus
        Exit Function
    End If
    If cboUserType.Text = "" Then
        MsgBox ("User Type required!")
        cboUserType.SetFocus
        Exit Function
    End If
    EntriesValid = True
End Function

Private Sub txtUserId_GotFocus()
    FocusMe (txtUserId)
End Sub

Private Sub txtUserId_LostFocus()
    If RecordExist("UserId", "dbo.Users", txtUserId.Text) Then
'        If LCase(User.UserId) <> LCase(txtUserId.Text) And LCase(User.UserType) <> "administrator" Then
'            'msgbox()
'        Else
            If MsgBox("User Id: " & txtUserId.Text & " already exist. Would you like to retrieve the information?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
                'call GetRecordset Function
                Dim rs As New ADODB.Recordset
                Set rs = GetRecordset("SELECT * FROM users WHERE userid = '" & txtUserId.Text & "'")
    
                If rs.State = adStateOpen Then
                    If Not rs.BOF Then 'if userid exists
                        txtUsername.Text = rs!UserName
                        txtPassword.Text = rs!Password
                        cboUserType.Text = rs!UserType
                        ValidateAccessLevel Me, "Update"
                    End If
                End If
            Else
                ClearEntries Me
                txtUserId.SetFocus
            End If
'        End If
    End If
End Sub

Private Sub txtUsername_GotFocus()
    FocusMe (txtUsername)
End Sub

Private Sub txtPassword_GotFocus()
    FocusMe (txtPassword)
End Sub



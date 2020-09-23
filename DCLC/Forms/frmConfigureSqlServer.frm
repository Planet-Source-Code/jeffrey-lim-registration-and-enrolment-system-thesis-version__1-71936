VERSION 5.00
Begin VB.Form frmConfigureSqlServer 
   Caption         =   "SQL Server Connection Configuration"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   6180
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      Picture         =   "frmConfigureSqlServer.frx":0000
      ScaleHeight     =   1065
      ScaleWidth      =   6150
      TabIndex        =   8
      Top             =   0
      Width           =   6180
   End
   Begin VB.Frame Frame1 
      Caption         =   " CONNECTION "
      Height          =   3495
      Left            =   45
      TabIndex        =   0
      Top             =   1350
      Width           =   6105
      Begin VB.OptionButton optSQLServer 
         Caption         =   "SQL Server"
         Height          =   285
         Left            =   3870
         TabIndex        =   11
         Top             =   1035
         Width           =   1410
      End
      Begin VB.OptionButton optWindows 
         Caption         =   "Windows"
         Height          =   240
         Left            =   2565
         TabIndex        =   10
         Top             =   1035
         Value           =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Height          =   525
         Left            =   4320
         TabIndex        =   4
         Top             =   2655
         Width           =   1620
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2565
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1935
         Width           =   2055
      End
      Begin VB.TextBox txtUserID 
         Height          =   285
         Left            =   2565
         TabIndex        =   2
         Top             =   1530
         Width           =   2055
      End
      Begin VB.TextBox txtServerName 
         Height          =   285
         Left            =   2565
         TabIndex        =   1
         Top             =   540
         Width           =   3315
      End
      Begin VB.Label Label2 
         Caption         =   "Authentication:"
         Height          =   255
         Left            =   300
         TabIndex        =   9
         Top             =   1035
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Password:"
         Height          =   255
         Left            =   765
         TabIndex        =   7
         Top             =   1845
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "User name:"
         Height          =   255
         Left            =   765
         TabIndex        =   6
         Top             =   1530
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Server Name or IP Address:"
         Height          =   255
         Left            =   300
         TabIndex        =   5
         Top             =   585
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmConfigureSqlServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sUser As String
Dim sComputer As String
Dim lpBuff As String * 1024

Private Sub Form_Load()
    txtServerName.Text = GetSetting(App.Title, "LOGIN", "SERVER", "(local)")
    optWindows_Click
End Sub

Private Sub cmdConnect_Click()
    'On Error GoTo errhandler
    If txtServerName.Text = "" Then
        Call MsgBox("Server Name is Needed.", vbOKOnly, "Server Name")
        Exit Sub
    End If
    
    ' Connect To Database
    If ConnectToSqlServer(txtServerName.Text, txtUserID.Text, txtPassword.Text) Then
        MsgBox ("Connection Succeeded")
        frmLogin.txtServer.Text = txtServerName.Text
        Unload Me
        frmLogin.show vbModal
    Else
        MsgBox ("Connection Failed")
        Exit Sub
    End If
errhandler:
End Sub

Private Sub optSQLServer_Click()
    txtUserID.Enabled = True
    txtPassword.Enabled = True
    txtUserID.Text = "sa"
End Sub

Private Sub optWindows_Click()
    txtUserID.Enabled = True
    txtPassword.Enabled = True
    'get machine name here
    'Get the Login User Name
    GetUserName lpBuff, Len(lpBuff)
    sUser = Left$(lpBuff, (InStr(1, lpBuff, vbNullChar)) - 1)
    lpBuff = ""
    
    'Get the Computer Name
    GetComputerName lpBuff, Len(lpBuff)
    sComputer = Left$(lpBuff, (InStr(1, lpBuff, vbNullChar)) - 1)
    lpBuff = ""
   
    'txtUserID.Text = sComputer & "\" & sUser
End Sub


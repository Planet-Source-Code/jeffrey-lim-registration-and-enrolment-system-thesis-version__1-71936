VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3765
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6000
   ControlBox      =   0   'False
   FillColor       =   &H000080FF&
   ForeColor       =   &H000080FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   3765
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   6510
      Top             =   2580
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1065
      Left            =   0
      TabIndex        =   0
      Top             =   2700
      Width           =   6075
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Version: 1.0.0.0000"
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
         Height          =   255
         Left            =   4275
         TabIndex        =   6
         Top             =   120
         Width           =   1980
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Licensee: Dr. CSLC"
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
         Height          =   255
         Left            =   4275
         TabIndex        =   5
         Top             =   360
         Width           =   1980
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Program © 2009 Dr. Carlos S. Lanting College"
         ForeColor       =   &H00636363&
         Height          =   285
         Left            =   900
         TabIndex        =   4
         Top             =   675
         Width           =   3345
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Author: Thesis Team"
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
         Height          =   255
         Left            =   4275
         TabIndex        =   3
         Top             =   660
         Width           =   1980
      End
      Begin VB.Image Image1 
         Height          =   750
         Left            =   60
         Picture         =   "frmSplash.frx":9748
         Top             =   120
         Width           =   750
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Dr. Carlos S. Lanting College"
         ForeColor       =   &H00636363&
         Height          =   285
         Left            =   900
         TabIndex        =   2
         Top             =   405
         Width           =   3300
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Dr.CSLC - RES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00636363&
         Height          =   285
         Left            =   900
         TabIndex        =   1
         Top             =   90
         Width           =   3315
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
  Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  Unload Me
End Sub

Private Sub Timer1_Timer()
  Unload Me
End Sub

Private Sub Form_Load()
  If App.PrevInstance = True Then
    MsgBox "System is Already in Run Mode ...", vbCritical, "System is Already Running ..."
    Unload Me
    Exit Sub
  End If
  ChDir (App.Path)
  'ShockwaveFlash1.Movie = App.Path & "\Flash\splash.swf" '-jil
End Sub


VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSy 
   Caption         =   "School Year"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmSy.frx":0000
   ScaleHeight     =   4200
   ScaleWidth      =   5565
   StartUpPosition =   3  'Windows Default
   Begin MSMask.MaskEdBox meStart 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   1890
      TabIndex        =   0
      Top             =   1395
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####"
      PromptChar      =   "_"
   End
   Begin MSDataGridLib.DataGrid dgSY 
      Height          =   1455
      Left            =   1485
      TabIndex        =   6
      Top             =   2070
      Width           =   3105
      _ExtentX        =   5477
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
      ScaleWidth      =   5505
      TabIndex        =   8
      Top             =   3705
      Width           =   5565
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "Close"
         Height          =   315
         Left            =   4335
         TabIndex        =   5
         Top             =   60
         Width           =   1035
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H000080FF&
         Caption         =   "Delete"
         Height          =   315
         Left            =   3255
         MaskColor       =   &H000080FF&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   60
         Width           =   1035
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   315
         Left            =   2175
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
      ScaleWidth      =   5565
      TabIndex        =   2
      Top             =   0
      Width           =   5565
      Begin VB.Image Image2 
         Height          =   375
         Left            =   5175
         Picture         =   "frmSy.frx":10380
         Top             =   0
         Width           =   390
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "School Year"
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
         Width           =   3135
      End
      Begin VB.Image Image1 
         Height          =   780
         Left            =   0
         Picture         =   "frmSy.frx":10830
         Top             =   60
         Width           =   915
      End
   End
   Begin MSMask.MaskEdBox meEnd 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   3150
      TabIndex        =   1
      Top             =   1395
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2745
      TabIndex        =   11
      Top             =   1440
      Width           =   285
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
      Left            =   720
      TabIndex        =   10
      Top             =   1485
      Width           =   1230
   End
End
Attribute VB_Name = "frmSy"
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
            RunSql ("DELETE FROM dbo.Sy WHERE SyId = '" + strSyId & "';")
            rsData.Requery
            cmdSave.Caption = "Save"
            strSyId = ""
            ClearEntries Me
            meStart.SetFocus
        End If
    Else
        MsgBox ("No record selected.")
    End If
End Sub

Private Sub cmdSave_Click()
    If EntriesValid Then
        If MsgBox("Do you want to " & cmdSave.Caption & " record?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
            strSql = "INSERT INTO dbo.Sy(Sy) VALUES('" & meStart & "-" & meEnd & "');"
            If cmdSave.Caption = "Update" Then
                strSql = "UPDATE dbo.Sy SET Sy = '" & meStart.Text & "-" & meEnd.Text & "' WHERE SyId = '" + strSyId & "';"
            End If
            RunSql (strSql)
            If rsData.State = adStateOpen Then
                rsData.Requery
            End If
            cmdSave.Caption = "Save"
            ClearEntries Me
            strSyId = ""
            meStart.SetFocus
        End If
    End If
End Sub

Private Sub dgSY_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    meStart.Text = Mid(rsData.Fields(1).Value, 1, 4)
    meEnd.Text = Mid(rsData.Fields(1).Value, 6, 4)
    strSyId = rsData.Fields(0).Value
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
    
    Set rsData = GetRecordset("SELECT syid AS ID, sy AS 'School Year' FROM sy ")
    If rsData.State = adStateOpen Then
        Set dgSY.DataSource = rsData
    End If
    ValidateAccessLevel Me, "Save"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsData.Close
End Sub

Public Function EntriesValid() As Boolean
    EntriesValid = False
    If Val(meStart.Text) < 2008 Then
        MsgBox ("Invalid School Year Start")
        meStart.SetFocus
        Exit Function
    End If
    If Val(meEnd.Text) <= Val(meStart.Text) Then
        MsgBox ("Invalid School Year End")
        meEnd.SetFocus
        Exit Function
    End If
    EntriesValid = True
End Function

Private Sub meEnd_GotFocus()
    FocusMe (meEnd)
End Sub

Private Sub meStart_GotFocus()
    FocusMe (meStart)
End Sub



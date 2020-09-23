VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmYearLevel 
   Caption         =   "YearLevel"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmYearLevel.frx":0000
   ScaleHeight     =   4125
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtYearLevel 
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
      Left            =   1530
      MaxLength       =   12
      TabIndex        =   0
      Top             =   1215
      Width           =   2220
   End
   Begin MSDataGridLib.DataGrid dgYearLevel 
      Height          =   1455
      Left            =   675
      TabIndex        =   4
      Top             =   1935
      Width           =   3420
      _ExtentX        =   6033
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
      ScaleWidth      =   4425
      TabIndex        =   7
      Top             =   3630
      Width           =   4485
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "Close"
         Height          =   315
         Left            =   3345
         TabIndex        =   3
         Top             =   60
         Width           =   1035
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H000080FF&
         Caption         =   "Delete"
         Height          =   315
         Left            =   2265
         MaskColor       =   &H000080FF&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   60
         Width           =   1035
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   315
         Left            =   1185
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
         TabIndex        =   8
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
      ScaleWidth      =   4485
      TabIndex        =   5
      Top             =   0
      Width           =   4485
      Begin VB.Image Image2 
         Height          =   375
         Left            =   4095
         Picture         =   "frmYearLevel.frx":10380
         Top             =   0
         Width           =   390
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Year Level"
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
         TabIndex        =   6
         Top             =   120
         Width           =   3135
      End
      Begin VB.Image Image1 
         Height          =   780
         Left            =   0
         Picture         =   "frmYearLevel.frx":10830
         Top             =   60
         Width           =   915
      End
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
      TabIndex        =   10
      Top             =   0
      Width           =   1230
   End
   Begin VB.Label Label1 
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
      Height          =   285
      Left            =   315
      TabIndex        =   9
      Top             =   1260
      Width           =   1230
   End
End
Attribute VB_Name = "frmYearLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsData As New ADODB.Recordset
Dim strYearLevelId As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If cmdSave.Caption = "Update" Then
        If MsgBox("Do you want to Delete the record?", vbYesNo + vbCritical, Me.Caption) = vbYes Then
            RunSql ("UPDATE dbo.YearLevel SET Deleted=1 WHERE YearLevelId = '" + strYearLevelId & "';")
            rsData.Requery
            ClearEntries Me
            strYearLevelId = ""
            txtYearLevel.SetFocus
        End If
    Else
        MsgBox ("No record selected.")
    End If
End Sub

Private Sub cmdSave_Click()
    If EntriesValid Then
        If MsgBox("Do you want to " & cmdSave.Caption & " record?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
            strSql = "INSERT INTO dbo.YearLevel(YearLevel, Deleted) " & _
                                    "VALUES('" & txtYearLevel.Text & "', " & _
                                           "'0');"
            If cmdSave.Caption = "Update" Then
                strSql = "UPDATE dbo.YearLevel SET YearLevel = '" & txtYearLevel.Text & "', " & _
                                                 "Deleted = '0' " & _
                                        "WHERE YearLevelId = '" & strYearLevelId & "';"
            End If
            RunSql (strSql)
            If rsData.State = adStateOpen Then
                rsData.Requery
            End If
            ClearEntries Me
            strYearLevelId = ""
            txtYearLevel.SetFocus
        End If
    End If
End Sub

Private Sub dgYearLevel_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    strYearLevelId = rsData.Fields(0).Value
    txtYearLevel = rsData.Fields(1).Value
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
    strSql = "SELECT YearLevelId AS ID, YearLevel FROM YearLevel WHERE deleted=0 ORDER BY YearLevel;"
    
    Set rsData = GetRecordset(strSql)
    If rsData.State = adStateOpen Then
        Set dgYearLevel.DataSource = rsData
    End If
    ValidateAccessLevel Me, "Save"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsData.Close
End Sub

Public Function EntriesValid() As Boolean
    EntriesValid = False
    If txtYearLevel.Text = "" Then
        MsgBox ("YearLevel required!")
        txtYearLevel.SetFocus
        Exit Function
    End If
    
    EntriesValid = True
End Function


Private Sub txtYearLevel_GotFocus()
    FocusMe (txtYearLevel)
End Sub




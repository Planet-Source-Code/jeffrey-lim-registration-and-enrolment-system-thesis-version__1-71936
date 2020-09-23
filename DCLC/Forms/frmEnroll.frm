VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmEnroll 
   Caption         =   "Enroll / Payment"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmEnroll.frx":0000
   ScaleHeight     =   533
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   759
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtYearLevelId 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   3150
      TabIndex        =   36
      Top             =   1080
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton cmdSearch 
      Height          =   315
      Left            =   2340
      Picture         =   "frmEnroll.frx":10380
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Search"
      Top             =   1035
      Width           =   360
   End
   Begin VB.ComboBox cboParticular 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmEnroll.frx":10755
      Left            =   7965
      List            =   "frmEnroll.frx":1076E
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   3015
      Width           =   3300
   End
   Begin MSMask.MaskEdBox meCredit 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   330
      Left            =   7965
      TabIndex        =   3
      Top             =   2655
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSComCtl2.DTPicker dtpTranDate 
      Height          =   330
      Left            =   7965
      TabIndex        =   2
      Top             =   2250
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   69795843
      CurrentDate     =   38207
   End
   Begin MSMask.MaskEdBox meAssessNo 
      Height          =   330
      Left            =   1125
      TabIndex        =   1
      Top             =   1035
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#######"
      PromptChar      =   "_"
   End
   Begin VB.PictureBox picFooter 
      Align           =   2  'Align Bottom
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   11325
      TabIndex        =   10
      Top             =   7500
      Width           =   11385
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "Close"
         Height          =   315
         Left            =   10185
         TabIndex        =   7
         Top             =   60
         Width           =   1035
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H000080FF&
         Caption         =   "Delete"
         Height          =   315
         Left            =   9105
         MaskColor       =   &H000080FF&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   60
         Width           =   1035
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   315
         Left            =   8025
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
      ScaleWidth      =   11385
      TabIndex        =   8
      Top             =   0
      Width           =   11385
      Begin MSMask.MaskEdBox meOrNo 
         Height          =   465
         Left            =   9270
         TabIndex        =   0
         Top             =   135
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   820
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         MaxLength       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "#######"
         PromptChar      =   "_"
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "RECEIPT#:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   7605
         TabIndex        =   13
         Top             =   225
         Width           =   1635
      End
      Begin VB.Image Image2 
         Height          =   375
         Left            =   11025
         Picture         =   "frmEnroll.frx":107BE
         Top             =   0
         Width           =   390
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Enroll / Payment"
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
         Width           =   4935
      End
      Begin VB.Image Image1 
         Height          =   780
         Left            =   0
         Picture         =   "frmEnroll.frx":10C6E
         Top             =   60
         Width           =   915
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmEnroll.frx":1145F
      Height          =   3960
      Left            =   45
      TabIndex        =   20
      Top             =   3465
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   6985
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      ForeColor       =   16576
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   6
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
            ColumnWidth     =   1514.986
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1514.986
         EndProperty
      EndProperty
   End
   Begin VB.Label lblYear 
      BackStyle       =   0  'Transparent
      Caption         =   "info"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   945
      TabIndex        =   34
      Top             =   2790
      Width           =   3750
   End
   Begin VB.Label lblCourse 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   945
      TabIndex        =   33
      Top             =   2475
      Width           =   5235
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   945
      TabIndex        =   32
      Top             =   2160
      Width           =   5460
   End
   Begin VB.Label lblStudentId 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   330
      Left            =   945
      TabIndex        =   31
      Top             =   1845
      Width           =   1815
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Particular:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   6840
      TabIndex        =   30
      Top             =   3015
      Width           =   1095
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amount:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   7155
      TabIndex        =   29
      Top             =   2655
      Width           =   780
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   6930
      TabIndex        =   28
      Top             =   2295
      Width           =   1005
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   270
      TabIndex        =   27
      Top             =   2205
      Width           =   645
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Outstanding"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   420
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   1995
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   2025
      TabIndex        =   25
      Top             =   0
      Width           =   1140
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Student"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   420
      Left            =   90
      TabIndex        =   24
      Top             =   1440
      Width           =   1365
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "info"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   1395
      TabIndex        =   23
      Top             =   1440
      Width           =   1140
   End
   Begin VB.Label lblBalance 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   8370
      TabIndex        =   22
      Top             =   1350
      Width           =   2625
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Assess # :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   90
      TabIndex        =   21
      Top             =   1080
      Width           =   825
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   9135
      TabIndex        =   19
      Top             =   945
      Width           =   1140
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   465
      Left            =   7110
      TabIndex        =   18
      Top             =   1845
      Width           =   1590
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   8595
      TabIndex        =   17
      Top             =   1845
      Width           =   1185
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Year:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   270
      TabIndex        =   16
      Top             =   2835
      Width           =   645
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Outstanding"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   420
      Left            =   7110
      TabIndex        =   15
      Top             =   945
      Width           =   1995
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   180
      TabIndex        =   14
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Id:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   360
      TabIndex        =   12
      Top             =   1890
      Width           =   555
   End
End
Attribute VB_Name = "frmEnroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsData As New ADODB.Recordset
Dim rsData2 As New ADODB.Recordset
Dim rsLedger As New ADODB.Recordset
Dim rsStudent As New ADODB.Recordset

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub meAssessNo_LostFocus()
    strSql = "SELECT StudentId FROM dbo.Enrolled " & _
            "WHERE AssessNo = '" & meAssessNo & "';"

    Set rsData = GetRecordset(strSql)
    If rsData.State = adStateOpen Then
        If rsData.RecordCount > 0 Then
            lblStudentId = rsData!StudentId
        Else
            MsgBox "Assessment # not found", vbInformation
        End If
    Else
        MsgBox "Assessment # not found", vbInformation
    End If
    Set rsData = Nothing

    lblStudentId_Change
End Sub

Private Sub meOrNo_GotFocus()
    ClearEntries Me
    lblStudentId = ""
    lblName = ""
    lblCourse = ""
    lblYear = ""
    ValidateAccessLevel Me, "Insert"
    meAssessNo.Enabled = True
    cmdSearch.Enabled = True
    meOrNo = GenerateNextPK("ReceiptNo")
    FocusMe (meAssessNo)
    Set DataGrid1.DataSource = Nothing
    DataGrid1.Refresh
End Sub

Private Sub meOrNo_LostFocus()
    strSql = "SELECT * " & _
             "FROM dbo.Ledger " & _
             "WHERE ReceiptNo = '" & meOrNo & "';"
     
    Set rsData2 = GetRecordset(strSql)
    If rsData2.State = adStateOpen Then
        If rsData2.RecordCount > 0 Then
            lblStudentId = rsData2!StudentId
            cboParticular = rsData2!Particular
            dtpTranDate = rsData2!TranDate
            meCredit = rsData2!Credit * -1
            cboParticular = rsData2!Particular
            meAssessNo.Mask = ""
            meAssessNo.Text = ""
            meAssessNo = Mid(rsData2!ReceiptNo, 1, 7)
            meAssessNo.Enabled = False
            cmdSearch.Enabled = False
            lblStudentId_Change
            ValidateAccessLevel Me, "Update"
        End If
    End If
    Set rsData2 = Nothing
End Sub


Private Sub cmdDelete_Click()
    If cmdSave.Caption = "Update" Then
        If MsgBox("Do you want to Delete the record?", vbYesNo + vbCritical, Me.Caption) = vbYes Then
            strSql = "DELETE FROM dbo.Ledger WHERE ReceiptNo = '" & meOrNo & "'; "
            'Execute SQL Command
            RunSql (strSql)
            
            ClearEntries Me
            ValidateAccessLevel Me, "Insert"
            meOrNo_GotFocus
            cmdSearch.SetFocus
        End If
    End If
End Sub

Private Sub cmdSave_Click()
    If EntriesValid Then
        If MsgBox("Do you want to " & cmdSave.Caption & " record?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
            strSql = "INSERT INTO dbo.Ledger(ReceiptNo, StudentId, SyId, YearLevelId, SemesterId, Credit, Particular, TranDate, PostedBy) " & _
                     "VALUES('" & meOrNo & "', '" & lblStudentId & "', " & SchoolInformation.CurrentSyId & ", " & _
                              txtYearLevelId & ", " & SchoolInformation.CurrentSemesterId & ", " & _
                              Val(meCredit) * -1 & ", " & _
                              "'" & Left(cboParticular, 25) & "', '" & Date & "', '" & User.UserId & "'); "
            If cmdSave.Caption = "Update" Then
                strSql = "UPDATE dbo.Ledger SET StudentId = '" & lblStudentId & "', " & _
                                               "SyId = " & SchoolInformation.CurrentSyId & ", " & _
                                               "YearLevelId = " & txtYearLevelId & ", " & _
                                               "SemesterId = " & SchoolInformation.CurrentSemesterId & ", " & _
                                               "Credit = " & Val(meCredit) * -1 & ", " & _
                                               "Particular = '" & Left(cboParticular, 25) & "', " & _
                                               "TranDate = '" & Date & "', " & _
                                               "PostedBy = '" & User.UserId & "'  " & _
                                    "WHERE ReceiptNo = '" & meOrNo & "'; "
            End If

            'Execute SQL Command
            RunSql (strSql)
            
            ClearEntries Me
            Call UpdatePK("ReceiptNo", cmdSave.Caption) 'Position this before ValidateAccessLevel.Inset
            ValidateAccessLevel Me, "Insert"
            meOrNo_GotFocus
            cmdSearch.SetFocus
        End If
    End If
End Sub

Private Sub cmdSearch_Click()
    frmSearchStudent.txtCallingForm = "frmEnroll"
    frmSearchStudent.Show vbModal
End Sub

Private Sub Form_Activate()
    MakeTransparent Me.hwnd, 255 'Unfade form
End Sub

Private Sub Form_Deactivate()
    MakeTransparent Me.hwnd, 200 'Fade Form
End Sub


Private Function EntriesValid() As Boolean
    EntriesValid = False
    If meCredit = "" Then
        MsgBox ("Payment amount required!")
        meCredit.SetFocus
        Exit Function
    End If
    If cboParticular = "" Then
        MsgBox ("Payment particular required!")
        cboParticular.SetFocus
        Exit Function
    End If

    EntriesValid = True
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set rsData = Nothing
    Set rsLedger = Nothing
    Set rsStudent = Nothing
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If rsLedger.State = adStateOpen Then
        If Not rsLedger.BOF And Not rsLedger.EOF Then
            
        End If
    End If
End Sub



Private Sub meStudentId_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub



Public Sub lblStudentId_Change()
    strSql = "SELECT *, (Lastname + ', ' + Firstname + ' ' + Middlename) AS Name " & _
             "FROM dbo.Students AS s " & _
                  "JOIN dbo.Courses AS c ON (s.CourseCode = c.CourseCode) " & _
                  "JOIN dbo.YearLevel AS yl ON (s.YearLevelId = yl.YearLevelId) " & _
                  "JOIN dbo.Semester AS sem ON (s.SemesterId = sem.SemesterId) " & _
            "WHERE s.StudentId = '" & lblStudentId & "';"
     
    Set rsData = GetRecordset(strSql)
    If rsData.State = adStateOpen Then
        If rsData.RecordCount > 0 Then
            lblName = rsData!Name
            lblCourse = rsData!CourseDesc
            lblYear = rsData!YearLevel
            
            txtYearLevelId = rsData!YearLevelId
        End If
    End If
    Set rsData = Nothing
    
    'Set Ledger
    strSql = "SELECT l1.LedgerId, l1.ReceiptNo, l1.TranDate, l1.Particular, l1.Debit, l1.Credit, SUM(l2.Debit + l2.Credit) AS Balance, l1.PostedBy, sy.Sy, yl.YearLevel, sem.Semester " & _
             "FROM dbo.Ledger AS l1 " & _
                  "INNER JOIN dbo.Sy AS sy ON (l1.SyId = sy.SyId) " & _
                  "INNER JOIN dbo.YearLevel AS yl ON (l1.YearLevelId = yl.YearLevelId) " & _
                  "INNER JOIN dbo.Semester AS sem ON (l1.SemesterId = sem.SemesterId) " & _
                  "INNER JOIN dbo.Ledger AS l2 ON (l1.TranDate >= l2.TranDate AND l1.LedgerId >= l2.LedgerId) " & _
            "WHERE l1.StudentId = '" & lblStudentId & "' " & _
            "GROUP BY l1.LedgerId, l1.ReceiptNo, l1.TranDate, l1.Particular, l1.Debit, l1.Credit, l1.PostedBy, sy.Sy, yl.YearLevel, sem.Semester " & _
            "ORDER BY TranDate;"
               
    Set rsLedger = GetRecordset(strSql)
    If rsLedger.State = adStateOpen Then
        If rsLedger.RecordCount > 0 Then
            'Get O/S Balance
            rsLedger.MoveLast
            lblBalance = Format(rsLedger!Balance, "###,##0.00")
            rsLedger.MoveFirst
        End If
    End If
    Set DataGrid1.DataSource = rsLedger
    DataGrid1.Columns(2).Width = 80
    DataGrid1.Columns(3).Width = 170
    DataGrid1.Columns(4).Alignment = dbgRight
    DataGrid1.Columns(4).NumberFormat = "###,##0.00"
    DataGrid1.Columns(5).NumberFormat = "###,##0.00"
    DataGrid1.Columns(6).NumberFormat = "###,##0.00"
    DataGrid1.Columns(5).Alignment = dbgRight
    DataGrid1.Columns(6).Alignment = dbgRight
End Sub



'Tab emulation
Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    EmulateEnter (KeyAscii) 'Tab emulation
End Sub

Private Sub dtpTranDate_KeyPress(KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub

Private Sub meCredit_KeyPress(KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub

Private Sub meOrNo_KeyPress(KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub


Private Sub meAssessNo_GotFocus()
    FocusMe (meAssessNo)
End Sub

Private Sub meAssessNo_KeyPress(KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub


Private Sub cboParticular_KeyPress(KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub

Private Sub meCredit_GotFocus()
    FocusMe (meCredit)
End Sub

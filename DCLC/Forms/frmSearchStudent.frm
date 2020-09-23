VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSearchStudent 
   Caption         =   "Search Student"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11070
   Icon            =   "frmSearchStudent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmSearchStudent.frx":08CA
   ScaleHeight     =   8040
   ScaleWidth      =   11070
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtAssessNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataField       =   "BirthDate"
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
      Left            =   9675
      MaxLength       =   20
      TabIndex        =   47
      Top             =   1200
      Width           =   1365
   End
   Begin VB.TextBox txtCallingForm 
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   285
      Left            =   9405
      TabIndex        =   46
      Text            =   "frmRegister"
      Top             =   4050
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox txtSemester 
      Appearance      =   0  'Flat
      DataField       =   "Semester"
      DataSource      =   "adoStudents"
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
      Left            =   1170
      MaxLength       =   35
      TabIndex        =   44
      Top             =   3915
      Width           =   3195
   End
   Begin VB.TextBox txtYear 
      Appearance      =   0  'Flat
      DataField       =   "YearLevel"
      DataSource      =   "adoStudents"
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
      Left            =   1170
      MaxLength       =   35
      TabIndex        =   43
      Top             =   3555
      Width           =   2835
   End
   Begin VB.TextBox txtCourse 
      Appearance      =   0  'Flat
      DataField       =   "CourseDesc"
      DataSource      =   "adoStudents"
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
      Left            =   1170
      MaxLength       =   35
      TabIndex        =   42
      Top             =   3195
      Width           =   6120
   End
   Begin VB.CheckBox chkGmrc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "GMRC"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   8550
      TabIndex        =   38
      Top             =   3375
      Width           =   1005
   End
   Begin VB.CheckBox chkBirthCertificate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Birth Certificate"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   9585
      TabIndex        =   37
      Top             =   3375
      Width           =   1410
   End
   Begin VB.CheckBox chkHsDiploma 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "High School Diploma"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   8550
      TabIndex        =   36
      Top             =   3690
      Width           =   1815
   End
   Begin VB.CheckBox chkForm137 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Form 137"
      DataField       =   "Form137"
      DataSource      =   "adoStudents"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   7470
      TabIndex        =   35
      Top             =   3330
      Width           =   3570
   End
   Begin VB.TextBox txtFather 
      Appearance      =   0  'Flat
      DataField       =   "Father"
      DataSource      =   "adoStudents"
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
      Left            =   990
      MaxLength       =   45
      TabIndex        =   23
      Top             =   2475
      Width           =   3795
   End
   Begin VB.TextBox txtMother 
      Appearance      =   0  'Flat
      DataField       =   "Mother"
      DataSource      =   "adoStudents"
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
      Left            =   5625
      MaxLength       =   45
      TabIndex        =   22
      Top             =   2475
      Width           =   3795
   End
   Begin VB.TextBox txtGender 
      Appearance      =   0  'Flat
      DataField       =   "Gender"
      DataSource      =   "adoStudents"
      Enabled         =   0   'False
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
      Left            =   7695
      MaxLength       =   20
      TabIndex        =   18
      Top             =   1710
      Width           =   510
   End
   Begin VB.TextBox txtReligion 
      Appearance      =   0  'Flat
      DataField       =   "BirthDate"
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
      Left            =   9135
      MaxLength       =   20
      TabIndex        =   17
      Top             =   1710
      Width           =   1905
   End
   Begin MSMask.MaskEdBox meStudentId 
      Height          =   465
      Left            =   8505
      TabIndex        =   16
      Top             =   225
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   820
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
      ScaleWidth      =   11010
      TabIndex        =   11
      Top             =   7545
      Width           =   11070
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "Previous"
         Height          =   315
         Left            =   7710
         TabIndex        =   14
         Top             =   60
         Width           =   1035
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         Height          =   315
         Left            =   8790
         MaskColor       =   &H000080FF&
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   60
         Width           =   1035
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "Ok"
         Height          =   315
         Left            =   9870
         TabIndex        =   12
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
         TabIndex        =   15
         Top             =   60
         Width           =   9915
      End
   End
   Begin VB.TextBox txtLastname 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
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
      Left            =   210
      MaxLength       =   35
      TabIndex        =   0
      Top             =   1200
      Width           =   3195
   End
   Begin VB.TextBox txtFirstname 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
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
      Left            =   3450
      MaxLength       =   35
      TabIndex        =   1
      Top             =   1200
      Width           =   3195
   End
   Begin VB.TextBox txtMiddlename 
      Appearance      =   0  'Flat
      DataField       =   "Middlename"
      DataSource      =   "adoStudents"
      Enabled         =   0   'False
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
      Left            =   6690
      MaxLength       =   35
      TabIndex        =   3
      Top             =   1200
      Width           =   2925
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      DataField       =   "Address"
      DataSource      =   "adoStudents"
      Enabled         =   0   'False
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
      Left            =   990
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1710
      Width           =   5940
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmSearchStudent.frx":10C4A
      Height          =   3195
      Left            =   45
      TabIndex        =   8
      Top             =   4320
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   5636
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
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
      ScaleWidth      =   11070
      TabIndex        =   9
      Top             =   0
      Width           =   11070
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "ID:"
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
         Left            =   7965
         TabIndex        =   45
         Top             =   270
         Width           =   465
      End
      Begin VB.Image Image1 
         Height          =   780
         Left            =   0
         Picture         =   "frmSearchStudent.frx":10C64
         Top             =   60
         Width           =   915
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Search Student"
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
         TabIndex        =   10
         Top             =   120
         Width           =   4305
      End
      Begin VB.Image Image2 
         Height          =   375
         Left            =   10665
         Picture         =   "frmSearchStudent.frx":11455
         Top             =   0
         Width           =   390
      End
   End
   Begin VB.CheckBox chkForm138 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Form 138"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   7470
      TabIndex        =   39
      Top             =   3645
      Width           =   3570
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Assessment #"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   9630
      TabIndex        =   48
      Top             =   945
      Width           =   1365
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "credentials"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8910
      TabIndex        =   41
      Top             =   2970
      Width           =   1410
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Submitted"
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
      Left            =   7335
      TabIndex        =   40
      Top             =   2925
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "course"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1350
      TabIndex        =   34
      Top             =   2925
      Width           =   1275
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Current"
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
      Left            =   90
      TabIndex        =   33
      Top             =   2880
      Width           =   1230
   End
   Begin VB.Label Label26 
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
      Left            =   270
      TabIndex        =   32
      Top             =   3915
      Width           =   1230
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
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
      Left            =   270
      TabIndex        =   31
      Top             =   3600
      Width           =   780
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Course"
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
      Left            =   270
      TabIndex        =   30
      Top             =   3240
      Width           =   1230
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "course"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1170
      TabIndex        =   29
      Top             =   45
      Width           =   1275
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Current"
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
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   1230
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "info"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1305
      TabIndex        =   27
      Top             =   2160
      Width           =   1275
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Parents"
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
      Left            =   135
      TabIndex        =   26
      Top             =   2115
      Width           =   1140
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Father:"
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
      Left            =   270
      TabIndex        =   25
      Top             =   2520
      Width           =   1230
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Mother:"
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
      Left            =   4905
      TabIndex        =   24
      Top             =   2520
      Width           =   1230
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
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
      Left            =   135
      TabIndex        =   21
      Top             =   1710
      Width           =   1230
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender:"
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
      Left            =   6975
      TabIndex        =   20
      Top             =   1800
      Width           =   690
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Birthdate:"
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
      Left            =   8295
      TabIndex        =   19
      Top             =   1755
      Width           =   780
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Student's No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   6570
      TabIndex        =   7
      Top             =   180
      Width           =   1650
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Lastname"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   135
      TabIndex        =   6
      Top             =   945
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Firstname"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3420
      TabIndex        =   5
      Top             =   945
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Miiddlename"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6660
      TabIndex        =   4
      Top             =   945
      Width           =   1095
   End
End
Attribute VB_Name = "frmSearchStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim adoStudents As New ADODB.Recordset
Dim adoAssessment As New ADODB.Recordset

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    cmdClose_Click
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    txtAssessNo.Text = ""
    If adoStudents.State = adStateOpen Then
        If Not adoStudents.BOF And Not adoStudents.EOF Then
            Set adoAssessment = GetRecordset("SELECT AssessNo FROM dbo.Enrolled " & _
                                             "WHERE StudentId = '" & adoStudents!StudentId & "' ORDER BY AssessNo DESC;")
            If adoAssessment.State = adStateOpen Then
                If adoAssessment.RecordCount > 0 Then
                    txtAssessNo.Text = adoAssessment!AssessNo
                End If
            End If
             Set adoAssessment = Nothing
        End If
    End If
End Sub

Private Sub Form_Activate()
    MakeTransparent Me.hwnd, 255 'Unfade form
End Sub

Private Sub Form_Deactivate()
    MakeTransparent Me.hwnd, 190 'Fade Form
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set adoStudents = Nothing
    If (Mid(meStudentId.Text, 1, 1) <> "_") And txtCallingForm.Text = "frmRegister" Then
        frmRegister.meStudentId = meStudentId
        frmRegister.meStudentId_LostFocus
    End If
    If (Mid(meStudentId.Text, 1, 1) <> "_") And txtCallingForm.Text = "frmLedger" Then
        frmLedger.meStudentId = meStudentId
        frmLedger.meStudentId_LostFocus
    End If
    If (Mid(meStudentId.Text, 1, 1) <> "_") And txtCallingForm.Text = "frmAssessSchedule" Then
        frmAssessSchedule.meStudentId = meStudentId
        frmAssessSchedule.meAssessNo.Mask = ""
        frmAssessSchedule.meAssessNo = Format(txtAssessNo, "0000000")
        frmAssessSchedule.meStudentId_LostFocus
    End If
    If (Mid(meStudentId.Text, 1, 1) <> "_") And txtCallingForm.Text = "frmAssess" Then
        frmAssess.meStudentId = meStudentId
        frmAssess.meStudentId_LostFocus
    End If
    If (Mid(meStudentId.Text, 1, 1) <> "_") And txtCallingForm.Text = "frmEnroll" Then
        If txtAssessNo.Text <> "" Then
            'frmEnroll.meAssessNo_LostFocus
            frmEnroll.meAssessNo.Mask = ""
            frmEnroll.meAssessNo = Format(txtAssessNo, "0000000")
        End If
        frmEnroll.lblStudentId = meStudentId
        frmEnroll.lblStudentId_Change
    End If
End Sub

Private Sub Form_Load()
    'On Error Resume Next
    strSql = "SELECT s.StudentId, Lastname, Firstname, Middlename, s.Address, Gender, BirthDate, LastSchoolAttended, " & _
                "CourseDesc, Semester, YearLevel, " & _
                "Father, Mother, " & _
                "Form137, Form138, Gmrc, BirthCertificate, HsDiploma " & _
         "FROM dbo.Students AS s " & _
         "INNER JOIN dbo.Courses AS c ON s.CourseCode = c.CourseCode " & _
         "INNER JOIN dbo.YearLevel AS yl ON s.YearLevelId = yl.YearLevelId " & _
         "INNER JOIN dbo.Semester AS sem ON s.SemesterId = sem.SemesterId " & _
         "INNER JOIN dbo.Parents AS p ON s.StudentId = p.StudentId " & _
         "INNER JOIN dbo.Credentials cr ON s.StudentId = cr.StudentId "
         
    Set adoStudents = GetRecordset(strSql)
    Set DataGrid1.DataSource = adoStudents
'    If adoStudents.State = adStateOpen Then
'                adoStudents.Requery
'            End If
'    txtLastname_LostFocus
End Sub



Private Sub cmdPrevious_Click()
    If adoStudents.State = adStateOpen Then
        If Not adoStudents.BOF Then
            adoStudents.MovePrevious
        Else
            adoStudents.MoveFirst
        End If
    End If
End Sub

Private Sub cmdNext_Click()
    If adoStudents.State = adStateOpen Then
        If Not adoStudents.EOF Then
            adoStudents.MoveNext
        Else
            adoStudents.MoveLast
        End If
  End If
End Sub

Function Search(Optional lStudNo As Boolean)
  Dim strPrevSQL As String
  
  If adoStudents.State = 1 Then Set adoStudents = Nothing
    
  strPrevSQL = strSql
  
  If lStudNo Then
    meStudentId.DataField = ""
    strSql = strSql & " WHERE s.StudentId = '" & meStudentId & "';"
    Debug.Print strSql
  Else
    meStudentId.DataField = "StudentId"
    If Trim(txtLastname) <> "" Then
        strSql = strSql & "WHERE LOWER([Lastname]) LIKE '%" & Trim(LCase(txtLastname)) & "%'"
        'strSQL = strSQL & "Where [Lastname] Like '" & Trim(UCase(txtLastname)) & "%'" 'removed as per panelist 2.21.5
    End If
    If Trim(txtLastname) <> "" And _
        Trim(txtFirstname) <> "" Then
        strSql = strSql & " AND "
    End If
    If Trim(txtLastname) = "" And _
        Trim(txtFirstname) <> "" Then
        strSql = strSql & " WHERE "
    End If
    If Trim(txtFirstname) <> "" Then
        'strSQL = strSQL & " [FirstName] Like '%" & Trim(UCase(txtFirstname)) & "%'" 'remove as per panelist 2.21.5
        strSql = strSql & " LOWER([FirstName]) LIKE '" & Trim(LCase(txtFirstname)) & "%'"
    End If
    strSql = strSql & " ORDER BY [LastName], [Firstname];"
  End If
  
  Set adoStudents = GetRecordset(strSql)
  
  Set DataGrid1.DataSource = adoStudents
  
  'Bind
  Set txtMiddlename.DataSource = adoStudents
  Set txtAddress.DataSource = adoStudents
  Set txtFather.DataSource = adoStudents
  Set txtMother.DataSource = adoStudents
  Set txtCourse.DataSource = adoStudents
    Set txtYear.DataSource = adoStudents
    Set txtSemester.DataSource = adoStudents
    Set chkForm137.DataSource = adoStudents
    Set chkForm138.DataSource = adoStudents
    Set chkGmrc.DataSource = adoStudents
    Set chkBirthCertificate.DataSource = adoStudents
    Set chkHsDiploma.DataSource = adoStudents
  
  If Not lStudNo Then
    meStudentId.DataField = "StudentId"
    If adoStudents.State = adStateOpen Then
        meStudentId.DataField = "StudentId"
        If adoStudents.RecordCount > 0 Then
            Set meStudentId.DataSource = adoStudents
        End If
    End If
  End If
  
  strSql = strPrevSQL

End Function


Private Sub meStudentId_LostFocus()
    Search (True)
    DataGrid1.SetFocus
End Sub

Private Sub txtLastname_LostFocus()
  Search (False)
End Sub

Private Sub txtFirstname_LostFocus()
  Search (False)
  DataGrid1.SetFocus
End Sub

'- UDF
Private Sub txtFirstname_Gotfocus()
  Call FocusMe(txtFirstname)
End Sub

Private Sub txtLastname_Gotfocus()
  Call FocusMe(txtLastname)
End Sub

Private Sub meStudentId_KeyPress(KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub


Private Sub txtLastname_KeyPress(KeyAscii As Integer)
  EmulateEnter (KeyAscii) 'Tab emulation
End Sub

Private Sub txtFirstname_KeyPress(KeyAscii As Integer)
  EmulateEnter (KeyAscii) 'Tab emulation
End Sub
'- eo: UDF


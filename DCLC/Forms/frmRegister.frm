VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRegister 
   Caption         =   "Register"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmRegister.frx":0000
   ScaleHeight     =   6255
   ScaleWidth      =   10035
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   90
      TabIndex        =   31
      Top             =   990
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   8281
      _Version        =   393216
      TabOrientation  =   2
      Tab             =   2
      TabHeight       =   520
      WordWrap        =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Credentials"
      TabPicture(0)   =   "frmRegister.frx":10380
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "chkForm137"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chkHsDiploma"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkBirthCertificate"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkGmrc"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkForm138"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label23"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label22"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Studies"
      TabPicture(1)   =   "frmRegister.frx":1039C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label20"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label21"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label24"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label25"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label26"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label27"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label28"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label29"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label30"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label31"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label32"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "dcCourses"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "dcYearLevel"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "dcSemester"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtLastSchoolAttended"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "cboStatus"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "General Info"
      TabPicture(2)   =   "frmRegister.frx":103B8
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label12"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label10"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label2"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label4"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label5"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label6"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label7"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label8"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label9"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label11"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Label13"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Label14"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label15"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Label16"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Label17"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Label18"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Label19"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "meStudentId"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "dtpBirthdate"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "txtLastname"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "txtFirstname"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "txtMiddlename"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "txtAddress"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "txtParentsAddress"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "txtFather"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "txtMother"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "txtFatherOccupation"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "txtMotherOccupation"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "cmdSearch"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "txtNationality"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "txtReligion"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "cboGender"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).ControlCount=   33
      Begin VB.ComboBox cboStatus 
         Height          =   315
         ItemData        =   "frmRegister.frx":103D4
         Left            =   -72705
         List            =   "frmRegister.frx":103F3
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   4005
         Width           =   1725
      End
      Begin VB.ComboBox cboGender 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmRegister.frx":10461
         Left            =   7605
         List            =   "frmRegister.frx":1046B
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   495
         Width           =   1230
      End
      Begin VB.TextBox txtReligion 
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
         Left            =   7605
         MaxLength       =   20
         TabIndex        =   8
         Top             =   1710
         Width           =   2175
      End
      Begin VB.TextBox txtNationality 
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
         Left            =   7605
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1305
         Width           =   2175
      End
      Begin VB.CommandButton cmdSearch 
         Height          =   315
         Left            =   3510
         Picture         =   "frmRegister.frx":10475
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Search"
         Top             =   540
         Width           =   360
      End
      Begin VB.TextBox txtMotherOccupation 
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
         Left            =   6300
         MaxLength       =   20
         TabIndex        =   12
         Top             =   3510
         Width           =   3480
      End
      Begin VB.TextBox txtFatherOccupation 
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
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   10
         Top             =   3510
         Width           =   3390
      End
      Begin VB.TextBox txtMother 
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
         Left            =   5940
         MaxLength       =   45
         TabIndex        =   11
         Top             =   3105
         Width           =   3840
      End
      Begin VB.TextBox txtFather 
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
         Left            =   1395
         MaxLength       =   45
         TabIndex        =   9
         Top             =   3105
         Width           =   3795
      End
      Begin VB.TextBox txtParentsAddress 
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
         MaxLength       =   100
         TabIndex        =   13
         Top             =   3915
         Width           =   7800
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
         Left            =   1890
         MaxLength       =   100
         TabIndex        =   4
         Top             =   2160
         Width           =   7800
      End
      Begin VB.TextBox txtMiddlename 
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
         Left            =   1890
         MaxLength       =   35
         TabIndex        =   3
         Top             =   1755
         Width           =   4470
      End
      Begin VB.TextBox txtFirstname 
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
         Left            =   1890
         MaxLength       =   35
         TabIndex        =   2
         Top             =   1350
         Width           =   4470
      End
      Begin VB.TextBox txtLastname 
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
         Left            =   1890
         MaxLength       =   35
         TabIndex        =   1
         Top             =   945
         Width           =   4470
      End
      Begin VB.CheckBox chkForm137 
         Caption         =   "Form 137 (Original Copy)"
         Height          =   375
         Left            =   -74010
         TabIndex        =   19
         Top             =   720
         Width           =   3480
      End
      Begin VB.CheckBox chkHsDiploma 
         Caption         =   "High School Diploma (Photocopy)"
         Height          =   375
         Left            =   -74010
         TabIndex        =   23
         Top             =   2655
         Width           =   3480
      End
      Begin VB.CheckBox chkBirthCertificate 
         Caption         =   "Birth Certificate (Photocopy)"
         Height          =   375
         Left            =   -74010
         TabIndex        =   22
         Top             =   2160
         Width           =   3480
      End
      Begin VB.CheckBox chkGmrc 
         Caption         =   "GMRC (Original Copy)"
         Height          =   375
         Left            =   -74010
         TabIndex        =   21
         Top             =   1665
         Width           =   3480
      End
      Begin VB.CheckBox chkForm138 
         Caption         =   "Form 138 (Original Copy)"
         Height          =   375
         Left            =   -74010
         TabIndex        =   20
         Top             =   1215
         Width           =   3480
      End
      Begin VB.TextBox txtLastSchoolAttended 
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
         Left            =   -72750
         MaxLength       =   30
         TabIndex        =   17
         Top             =   2880
         Width           =   4470
      End
      Begin MSDataListLib.DataCombo dcSemester 
         Height          =   315
         Left            =   -72750
         TabIndex        =   16
         Top             =   1530
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcYearLevel 
         Height          =   315
         Left            =   -72750
         TabIndex        =   15
         Top             =   1080
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcCourses 
         Height          =   315
         Left            =   -72750
         TabIndex        =   14
         Top             =   630
         Width           =   5610
         _ExtentX        =   9895
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker dtpBirthdate 
         Height          =   330
         Left            =   7605
         TabIndex        =   6
         Top             =   855
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "MMM-dd-yyyy"
         Format          =   69861379
         CurrentDate     =   38207
      End
      Begin MSMask.MaskEdBox meStudentId 
         Height          =   330
         Left            =   1890
         TabIndex        =   0
         Top             =   540
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##-#######"
         PromptChar      =   "_"
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "status"
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
         Left            =   -72525
         TabIndex        =   63
         Top             =   3600
         Width           =   1275
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Student"
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
         Left            =   -74280
         TabIndex        =   62
         Top             =   3555
         Width           =   1725
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Category:"
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
         Left            =   -74145
         TabIndex        =   61
         Top             =   4050
         Width           =   1230
      End
      Begin VB.Label Label29 
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
         Left            =   -74145
         TabIndex        =   59
         Top             =   2925
         Width           =   1230
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
         Height          =   420
         Left            =   -74145
         TabIndex        =   58
         Top             =   630
         Width           =   1230
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Year Level"
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
         Left            =   -74145
         TabIndex        =   57
         Top             =   1080
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
         Left            =   -74145
         TabIndex        =   56
         Top             =   1575
         Width           =   1230
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Last School"
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
         Left            =   -74280
         TabIndex        =   55
         Top             =   2250
         Width           =   1725
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "attended"
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
         Left            =   -72525
         TabIndex        =   54
         Top             =   2295
         Width           =   1275
      End
      Begin VB.Label Label23 
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
         Left            =   -74235
         TabIndex        =   53
         Top             =   180
         Width           =   1455
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
         Left            =   -72660
         TabIndex        =   52
         Top             =   225
         Width           =   1410
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
         Left            =   -74325
         TabIndex        =   51
         Top             =   135
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
         Left            =   -73155
         TabIndex        =   50
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label Label19 
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
         Height          =   240
         Left            =   720
         TabIndex        =   49
         Top             =   3915
         Width           =   1230
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation:"
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
         Left            =   5265
         TabIndex        =   48
         Top             =   3555
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
         Left            =   5265
         TabIndex        =   47
         Top             =   3150
         Width           =   1230
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation:"
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
         TabIndex        =   46
         Top             =   3510
         Width           =   1230
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
         Left            =   720
         TabIndex        =   45
         Top             =   3150
         Width           =   1230
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
         Left            =   540
         TabIndex        =   44
         Top             =   2700
         Width           =   1140
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
         Left            =   1710
         TabIndex        =   43
         Top             =   2745
         Width           =   1275
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Religion:"
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
         Left            =   6615
         TabIndex        =   42
         Top             =   1755
         Width           =   1230
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Nationality:"
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
         Left            =   6615
         TabIndex        =   41
         Top             =   1350
         Width           =   1230
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
         Left            =   6615
         TabIndex        =   40
         Top             =   900
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
         Left            =   6615
         TabIndex        =   39
         Top             =   540
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
         Height          =   240
         Left            =   705
         TabIndex        =   38
         Top             =   2205
         Width           =   1230
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Middlename:"
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
         Left            =   705
         TabIndex        =   37
         Top             =   1800
         Width           =   1230
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Firstname:"
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
         Left            =   705
         TabIndex        =   36
         Top             =   1395
         Width           =   1230
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Lastname:"
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
         TabIndex        =   35
         Top             =   990
         Width           =   1230
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Personal"
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
         Left            =   540
         TabIndex        =   34
         Top             =   90
         Width           =   1275
      End
      Begin VB.Label Label12 
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
         Left            =   1890
         TabIndex        =   33
         Top             =   135
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Student #:"
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
         Left            =   705
         TabIndex        =   32
         Top             =   585
         Width           =   1230
      End
   End
   Begin VB.PictureBox picFooter 
      Align           =   2  'Align Bottom
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   9975
      TabIndex        =   29
      Top             =   5760
      Width           =   10035
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "Close"
         Height          =   315
         Left            =   8880
         TabIndex        =   26
         Top             =   60
         Width           =   1035
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H000080FF&
         Caption         =   "Delete"
         Height          =   315
         Left            =   7800
         MaskColor       =   &H000080FF&
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   60
         Width           =   1035
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   315
         Left            =   6720
         TabIndex        =   24
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
         TabIndex        =   30
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
      ScaleWidth      =   10035
      TabIndex        =   27
      Top             =   0
      Width           =   10035
      Begin VB.Image Image2 
         Height          =   375
         Left            =   9660
         Picture         =   "frmRegister.frx":1084A
         Top             =   0
         Width           =   390
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Register"
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
         TabIndex        =   28
         Top             =   120
         Width           =   3135
      End
      Begin VB.Image Image1 
         Height          =   780
         Left            =   0
         Picture         =   "frmRegister.frx":10CFA
         Top             =   60
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsData As New ADODB.Recordset



Private Sub cboGender_KeyPress(KeyAscii As Integer)
    EmulateEnter KeyAscii
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdDelete_Click()
    If cmdSave.Caption = "Update" Then
        If MsgBox("Do you want to Delete the record?", vbYesNo + vbCritical, Me.Caption) = vbYes Then
            strSql = "UPDATE dbo.Students SET  Status = 'Deleted' " & _
                        "WHERE StudentId = '" & meStudentId.Text & "'; "
            'Execute SQL Command
            RunSql (strSql)
            
            If rsData.State = adStateOpen Then
                rsData.Requery
            End If
            meStudentId.SetFocus
            txtLastname.SetFocus
            SSTab1.Tab = 2
        End If
    End If
End Sub

Private Sub cmdSave_Click()
    'UpdatePK ("StudentNo")
    If EntriesValid Then
        If MsgBox("Do you want to " & cmdSave.Caption & " record?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
            strSql = "INSERT INTO dbo.Students(StudentId, Lastname, Firstname, Middlename, Address, " & _
                                              "Gender, Nationality, Religion, " & _
                                              "CourseCode, YearLevelId, SemesterId, " & _
                                              "LastSchoolAttended, BirthDate, DateEnrolled, Status) " & _
                                    "VALUES('" & meStudentId.Text & "', " & _
                                           "'" & txtFirstname.Text & "', " & _
                                           "'" & txtLastname.Text & "', " & _
                                           "'" & txtMiddlename.Text & "', " & _
                                           "'" & txtAddress.Text & "', " & _
                                           "'" & cboGender.Text & "', " & _
                                           "'" & txtNationality.Text & "', " & _
                                           "'" & txtReligion.Text & "', " & _
                                           "'" & dcCourses.BoundText & "', " & _
                                           "'" & dcYearLevel.BoundText & "', " & _
                                           "'" & dcSemester.BoundText & "', " & _
                                           "'" & txtLastSchoolAttended.Text & "', " & _
                                           "'" & dtpBirthdate.Value & "', " & _
                                           "'" & Date & "', " & _
                                           "'" & cboStatus.Text & "'); "
            strSql = strSql + "INSERT INTO dbo.Parents(StudentId, Father, FatherOccupation, Mother, MotherOccupation, Address) " & _
                                       "VALUES('" & meStudentId.Text & "', " & _
                                              "'" & txtFather.Text & "', " & _
                                              "'" & txtFatherOccupation.Text & "', " & _
                                              "'" & txtMother.Text & "', " & _
                                              "'" & txtMotherOccupation.Text & "', " & _
                                              "'" & txtParentsAddress.Text & "'); "
            strSql = strSql + "INSERT INTO dbo.Credentials(StudentId, Form137, Form138, Gmrc, BirthCertificate, HsDiploma) " & _
                                       "VALUES('" & meStudentId.Text & "', " & _
                                              "'" & chkForm137.Value & "', " & _
                                              "'" & chkForm138.Value & "', " & _
                                              "'" & chkGmrc.Value & "', " & _
                                              "'" & chkBirthCertificate.Value & "', " & _
                                              "'" & chkHsDiploma.Value & "'); "
            If cmdSave.Caption = "Update" Then
                strSql = "UPDATE dbo.Students SET  Lastname = '" & txtLastname.Text & "', " & _
                                                  "Firstname = '" & txtFirstname.Text & "', " & _
                                                  "Middlename = '" & txtMiddlename.Text & "', " & _
                                                  "Address = '" & txtAddress.Text & "', " & _
                                                  "Gender = '" & cboGender.Text & "', " & _
                                                  "Birthdate = '" & dtpBirthdate.Value & "', " & _
                                                  "Nationality = '" & txtNationality.Text & "', " & _
                                                  "Religion = '" & txtReligion.Text & "', " & _
                                                  "CourseCode = '" & dcCourses.BoundText & "', " & _
                                                  "YearLevelId = '" & dcYearLevel.BoundText & "', " & _
                                                  "SemesterId = '" & dcSemester.BoundText & "', " & _
                                                  "LastSchoolAttended = '" & txtLastSchoolAttended & "', " & _
                                                  "DateEnrolled = '" & Date & "', " & _
                                                  "Status = '" & cboStatus.Text & "' " & _
                        "WHERE StudentId = '" & meStudentId.Text & "'; "
                strSql = strSql + "UPDATE dbo.Parents SET Father = '" & txtFather.Text & "', " & _
                                                        "FatherOccupation = '" & txtFatherOccupation.Text & "', " & _
                                                        "Mother = '" & txtMother.Text & "', " & _
                                                        "MotherOccupation = '" & txtMotherOccupation.Text & "', " & _
                                                        "Address = '" & txtParentsAddress.Text & "' " & _
                                            "WHERE StudentId = '" & meStudentId.Text & "'; "
                strSql = strSql + "UPDATE dbo.Credentials SET Form137 = '" & chkForm137.Value & "', " & _
                                                             "Form138 = '" & chkForm137.Value & "', " & _
                                                             "Gmrc = '" & chkForm137.Value & "', " & _
                                                             "BirthCertificate = '" & chkForm137.Value & "', " & _
                                                             "HsDiploma = '" & chkForm137.Value & "' " & _
                                       "WHERE StudentId = '" & meStudentId.Text & "'; "

            End If

            'Execute SQL Command
            RunSql (strSql)
            
            If rsData.State = adStateOpen Then
                rsData.Requery
            End If
            ClearEntries Me
            Call UpdatePK("StudentNo", cmdSave.Caption) 'Position this before ValidateAccessLevel.Inset
            ValidateAccessLevel Me, "Insert"
            meStudentId.SetFocus
            txtLastname.SetFocus
            SSTab1.Tab = 2
        End If
    End If
End Sub

Private Sub cmdSearch_Click()
    frmSearchStudent.txtCallingForm = "frmRegister"
    frmSearchStudent.Show vbModal
End Sub



Private Sub dtpBirthdate_KeyPress(KeyAscii As Integer)
EmulateEnter KeyAscii
End Sub

Private Sub Form_Load()
    SSTab1.Tab = 2
    BindDataCombo "SELECT * FROM dbo.Courses", "CourseDesc", dcCourses, "CourseCode", True 'Bind Courses
    BindDataCombo "SELECT * FROM dbo.YearLevel", "YearLevel", dcYearLevel, "YearLevelId", True 'Bind School Year
    BindDataCombo "SELECT * FROM dbo.Semester", "Semester", dcSemester, "SemesterId", True 'Bind Semester
End Sub

Private Sub Form_Activate()
'    MakeTransparent Me.hwnd, 255 'Unfade form
End Sub

Private Sub Form_Deactivate()
'    MakeTransparent Me.hwnd, 200 'Fade Form
End Sub


Private Function EntriesValid() As Boolean
    EntriesValid = False
    If txtFirstname.Text = "" Then
        MsgBox ("Firstname required!")
        SSTab1.Tab = 2
        txtFirstname.SetFocus
        Exit Function
    End If
    If txtLastname.Text = "" Then
        MsgBox ("Lastname required!")
        SSTab1.Tab = 2
        txtLastname.SetFocus
        Exit Function
    End If
    If txtMiddlename.Text = "" Then
        MsgBox ("Middlename required!")
        SSTab1.Tab = 2
        txtMiddlename.SetFocus
        Exit Function
    End If
    If txtAddress.Text = "" Then
        MsgBox ("Address required!")
        SSTab1.Tab = 2
        txtAddress.SetFocus
        Exit Function
    End If
    If cboGender.Text = "" Then
        MsgBox ("Gender required!")
        SSTab1.Tab = 2
        cboGender.SetFocus
        Exit Function
    End If
    If dtpBirthdate.Value = "" Then
        MsgBox ("Birthdate required!")
        SSTab1.Tab = 2
        dtpBirthdate.SetFocus
        Exit Function
    End If
    If txtNationality.Text = "" Then
        MsgBox ("Nationality required!")
        SSTab1.Tab = 2
        txtNationality.SetFocus
        Exit Function
    End If
    If txtReligion.Text = "" Then
        MsgBox ("Religion required!")
        SSTab1.Tab = 2
        txtReligion.SetFocus
        Exit Function
    End If
    If txtFather.Text = "" Then
        MsgBox ("Father required!")
        SSTab1.Tab = 2
        txtFather.SetFocus
        Exit Function
    End If
    If txtFatherOccupation.Text = "" Then
        MsgBox ("Father's Occupation required!")
        SSTab1.Tab = 2
        txtFatherOccupation.SetFocus
        Exit Function
    End If
    If txtMother.Text = "" Then
        MsgBox ("Mother required!")
        SSTab1.Tab = 2
        txtMother.SetFocus
        Exit Function
    End If
    If txtMotherOccupation.Text = "" Then
        MsgBox ("Mother's Occupation required!")
        SSTab1.Tab = 2
        txtMotherOccupation.SetFocus
        Exit Function
    End If
    If txtParentsAddress.Text = "" Then
        MsgBox ("Parents' Address required!")
        SSTab1.Tab = 2
        txtParentsAddress.SetFocus
        Exit Function
    End If
    If dcCourses.Tag = "" Then
        MsgBox ("Course required!")
        SSTab1.Tab = 1
        dcCourses.SetFocus
        Exit Function
    End If
    If dcYearLevel.Tag = "" Then
        MsgBox ("Year Level required!")
        SSTab1.Tab = 1
        dcYearLevel.SetFocus
        Exit Function
    End If
    If dcSemester.Tag = "" Then
        MsgBox ("Semester required!")
        SSTab1.Tab = 1
        dcSemester.SetFocus
        Exit Function
    End If
    If txtLastSchoolAttended.Text = "" Then
        MsgBox ("Last School Attended required!")
        SSTab1.Tab = 1
        txtLastSchoolAttended.SetFocus
        Exit Function
    End If
    If cboStatus.Text = "" Then
        MsgBox ("Student category status required!")
        SSTab1.Tab = 1
        cboStatus.SetFocus
        Exit Function
    End If
    EntriesValid = True
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set rsData = Nothing
End Sub

Private Sub meStudentId_GotFocus()
    ValidateAccessLevel Me, "Insert"
    ClearEntries Me
    meStudentId.Text = GenerateNextPK("StudentNo")
    FocusMe (meStudentId)
End Sub

Private Sub meStudentId_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Public Sub meStudentId_LostFocus()
    strSql = "SELECT *, p.Address AS ParentsAddress " & _
             "FROM dbo.Students AS s " & _
                  "JOIN dbo.Courses AS c ON (s.CourseCode = c.CourseCode) " & _
                  "JOIN dbo.YearLevel AS yl ON (s.YearLevelId = yl.YearLevelId) " & _
                  "JOIN dbo.Semester AS sem ON (s.SemesterId = sem.SemesterId) " & _
                  "JOIN dbo.Parents AS p ON (s.StudentId = p.StudentId) " & _
                  "JOIN dbo.Credentials AS cr ON (s.StudentId = cr.StudentId) " & _
            "WHERE s.StudentId = '" & meStudentId.Text & "';"
     
    Set rsData = GetRecordset(strSql)
    If rsData.State = adStateOpen Then
        If rsData.RecordCount > 0 Then
            txtLastname.Text = rsData!Lastname
            txtFirstname.Text = rsData!Firstname
            txtMiddlename.Text = rsData!Middlename
            txtAddress.Text = rsData!Address
            cboGender.Text = rsData!Gender
            dtpBirthdate = rsData!BirthDate
            txtNationality.Text = rsData!Nationality
            txtReligion.Text = rsData!Religion
            txtFather.Text = rsData!Father
            txtFatherOccupation.Text = rsData!FatherOccupation
            txtMother.Text = rsData!Mother
            txtMotherOccupation.Text = rsData!MotherOccupation
            txtParentsAddress.Text = rsData!ParentsAddress
            dcCourses.BoundText = rsData!CourseCode
            dcYearLevel.BoundText = rsData!YearLevelId
            dcSemester.BoundText = rsData!SemesterId
            txtLastSchoolAttended.Text = rsData!LastSchoolAttended
            cboStatus.Text = rsData!Status
            chkForm137.Value = IIf(rsData!Form137, 1, 0)
            chkForm138.Value = IIf(rsData!Form138, 1, 0)
            chkGmrc.Value = IIf(rsData!Gmrc, 1, 0)
            chkBirthCertificate.Value = IIf(rsData!BirthCertificate, 1, 0)
            chkHsDiploma.Value = IIf(rsData!HsDiploma, 1, 0)
            ValidateAccessLevel Me, "Update"
        End If
    End If
    
End Sub

'Tab emulation
Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtFather_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtFatherOccupation_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtFirstname_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtLastname_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtLastSchoolAttended_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtMiddlename_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtMother_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtMotherOccupation_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtNationality_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtParentsAddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtReligion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

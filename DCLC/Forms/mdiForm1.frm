VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Dr. Carlos Lanting College - Registration & Enrollment Systems"
   ClientHeight    =   7980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11355
   Icon            =   "mdiForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imgBoxX 
      Left            =   9450
      Top             =   675
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   138
      ImageHeight     =   128
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":22952
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":238CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":245F2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgSearch0 
      Left            =   8235
      Top             =   675
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":25343
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgFooter 
      Left            =   10080
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   759
      ImageHeight     =   49
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":256FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":266EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":27672
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgHeader 
      Left            =   9450
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   759
      ImageHeight     =   119
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":28609
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":2EAA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":34A0C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgBox1 
      Left            =   8820
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   138
      ImageHeight     =   128
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":3AE3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":3CF85
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":3EB8A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgBox0 
      Left            =   8235
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   138
      ImageHeight     =   128
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":4074A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":4205A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":43582
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrResize 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7740
      Top             =   315
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   7695
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   12
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1411
            MinWidth        =   1411
            Picture         =   "mdiForm1.frx":44B9A
            Text            =   "User:"
            TextSave        =   "User:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   441
            MinWidth        =   441
            Picture         =   "mdiForm1.frx":44FAB
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "Log-in:"
            TextSave        =   "Log-in:"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1765
            MinWidth        =   1765
            TextSave        =   "3:29 AM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
            Picture         =   "mdiForm1.frx":45345
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "Date:"
            TextSave        =   "Date:"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1834
            MinWidth        =   1834
            TextSave        =   "3/7/2009"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1058
            MinWidth        =   1058
            Picture         =   "mdiForm1.frx":456DF
            Text            =   "SY:"
            TextSave        =   "SY:"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1834
            MinWidth        =   1834
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1236
            MinWidth        =   1236
            Picture         =   "mdiForm1.frx":45A96
            Text            =   "Sem:"
            TextSave        =   "Sem:"
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picWall 
      Align           =   1  'Align Top
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   7620
      Left            =   0
      Picture         =   "mdiForm1.frx":45EB4
      ScaleHeight     =   7620
      ScaleWidth      =   11355
      TabIndex        =   1
      Top             =   0
      Width           =   11355
      Begin VB.PictureBox picSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   10665
         Picture         =   "mdiForm1.frx":57AA8
         ScaleHeight     =   645
         ScaleWidth      =   675
         TabIndex        =   7
         Top             =   6075
         Width           =   675
      End
      Begin VB.PictureBox picBox3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1995
         Left            =   8010
         Picture         =   "mdiForm1.frx":57E50
         ScaleHeight     =   1995
         ScaleWidth      =   2070
         TabIndex        =   6
         Top             =   3015
         Width           =   2070
      End
      Begin VB.PictureBox picBox2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1995
         Left            =   4680
         Picture         =   "mdiForm1.frx":59458
         ScaleHeight     =   1995
         ScaleWidth      =   2070
         TabIndex        =   5
         Top             =   2070
         Width           =   2070
      End
      Begin VB.PictureBox picBox1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1995
         Left            =   1215
         Picture         =   "mdiForm1.frx":5A970
         ScaleHeight     =   1995
         ScaleWidth      =   2070
         TabIndex        =   4
         Top             =   3015
         Width           =   2070
      End
      Begin VB.PictureBox picHeader 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1770
         Left            =   0
         Picture         =   "mdiForm1.frx":5C270
         ScaleHeight     =   1770
         ScaleWidth      =   11385
         TabIndex        =   3
         Top             =   0
         Width           =   11385
         Begin MSComctlLib.ImageList imgSearch1 
            Left            =   8820
            Top             =   675
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   40
            ImageHeight     =   40
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiForm1.frx":626FC
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin VB.PictureBox picFooter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   0
         Picture         =   "mdiForm1.frx":62BBB
         ScaleHeight     =   735
         ScaleWidth      =   11385
         TabIndex        =   2
         Top             =   6750
         Width           =   11385
      End
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   35
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":63B9F
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":64879
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":656CB
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":6651D
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":66DF7
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":676D1
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":67FAB
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":68975
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":6924F
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":69569
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":69E43
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":6A71D
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":6AFF7
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":6B311
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":6BBEB
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":6C4C5
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":6CD9F
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":6D679
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":6DF53
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":6E82D
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":6F107
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":6F9E1
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":702BB
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":70B95
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":7146F
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":71D49
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":72623
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":72EFD
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":737D7
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":740B1
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":74967
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":75241
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":75693
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":75AE5
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiForm1.frx":78297
            Key             =   "IMG35"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuRegister 
      Caption         =   "Register"
   End
   Begin VB.Menu mnuAssessment 
      Caption         =   "Assessment"
   End
   Begin VB.Menu mnuAccounting 
      Caption         =   "Accounting"
      Begin VB.Menu mnuAccountSidebar 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Acctg|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:000|Gradient}"
      End
      Begin VB.Menu mnuEnroll 
         Caption         =   "Enroll / Payment"
      End
      Begin VB.Menu mnuFees 
         Caption         =   "Fee Schedule"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports"
      Begin VB.Menu mnuReportsSidebar 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Reports|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:000|Gradient}"
      End
      Begin VB.Menu mnuStudentList 
         Caption         =   "List of Students"
      End
      Begin VB.Menu mnuPopulation 
         Caption         =   "Population"
      End
      Begin VB.Menu mnuFeesAndSchedules 
         Caption         =   "Fees Breakdown && Schedules"
      End
      Begin VB.Menu mnuLedger 
         Caption         =   "Student's Ledger"
      End
   End
   Begin VB.Menu mnuMaintenance 
      Caption         =   "Maintenance"
      Begin VB.Menu mnuSeparator1 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Maintenance|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:000|Gradient}"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSetup 
         Caption         =   "Setup"
         Begin VB.Menu mnuUsers 
            Caption         =   "Users"
         End
         Begin VB.Menu mnuSchoolInfo 
            Caption         =   "School Information && Current School Year"
         End
      End
      Begin VB.Menu mnuSchoolYear 
         Caption         =   "School Year"
      End
      Begin VB.Menu mnuEducationLevel 
         Caption         =   "Education Level"
      End
      Begin VB.Menu mnuSemester 
         Caption         =   "Semester"
      End
      Begin VB.Menu mnuCourse 
         Caption         =   "Courses"
      End
      Begin VB.Menu mnuDependency 
         Caption         =   "System Dependency"
         Begin VB.Menu mnuSchedulingSystem 
            Caption         =   "Scheduling System"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Activate()
    'If end_app = True Then End
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    ReleaseMenus hwnd
    'Terminate the entire application
    End
End Sub

Private Sub MDIForm_Load()
    'Set Gradient Menus
    SetMenus hwnd, SmallImages
    
    '*********************************************************************
    dbServer = GetSetting(App.Title, "LOGIN", "SERVER", "(local)")
    '\HKEY_CURRENT_USER\Software\VB and VBA Program Settings\Sample\LOGIN [SERVER]
    'retrieves value stored in the registry (local) stands for the local machine used as server [default server].
    'you can also use the computer name of your local machine or the computer name within the network being used as the server.
    '*********************************************************************
    
    If ConnectToSqlServer(dbServer, "", "") Then  'open database connection
        'MsgBox ("SQL Server Connection succeeded")
    Else
        Call MsgBox("Pls click OK to configure SQL Server Connection.", vbCritical, "Database Connection Error. ")
        frmConfigureSqlServer.Show vbModal
    End If
    
    Me.Show
    frmSplash.Show vbModal

    frmLogin.txtServer.Text = dbServer
    frmLogin.Show vbModal

    '-- Determine Access Depth Here
    mnuAssessment.Enabled = LCase(User.UserType) <> "registrar"
    mnuAccounting.Enabled = LCase(User.UserType) <> "registrar"
    mnuLedger.Enabled = LCase(User.UserType) <> "registrar"

    mnuSchoolInfo.Enabled = (LCase(User.UserType) = "administrator")
    mnuSchoolYear.Enabled = (LCase(User.UserType) = "administrator")
    mnuEducationLevel.Enabled = (LCase(User.UserType) = "administrator")
    mnuSemester.Enabled = (LCase(User.UserType) = "administrator")
    mnuCourse.Enabled = (LCase(User.UserType) = "administrator")
    
    picWall_Move
    '-- eo:

'    'For testing Only
'    User.UserId = "admin"
'    User.UserName = "Administrator"
'    User.UserType = "administrator"
'    SchoolInformation.CurrentSyId = 1
'    SchoolInformation.CurrentSemesterId = 1
'    SchoolInformation.Semester = "1st Semester"
'    SchoolInformation.Sy = "2008-2009"
'    With MDIForm1.StatusBar1.Panels
'        .Item(10).Text = "2008-2009"
'        .Item(12).Text = "1st Semester"
'    End With
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("Exit Application?", vbExclamation + vbYesNo, "DCLC-RES") = vbNo Then
        Cancel = 1
    End If
End Sub



Private Sub mnuC_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuCalc_Click()
    On Error GoTo err
    Shell "calc.exe", vbNormalFocus
    Exit Sub
err:
    MsgBox "You don't have a Calculator installed in your computer.", vbExclamation, "DCLC-RES Thesis Version"
End Sub

Private Sub mnuTH_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuTV_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuArrangeIcons_Click()
    MDIForm1.Arrange vbArrangeIcons
End Sub

Private Sub mnuCascade_Click()
    MDIForm1.Arrange vbCascade
End Sub

Private Sub mnuAssessment_Click()
    frmAssess.Show
End Sub

Private Sub mnuCourse_Click()
    frmCourses.Show
End Sub

Private Sub mnuEducationLevel_Click()
    frmYearLevel.Show
End Sub

Private Sub mnuEnroll_Click()
    frmEnroll.Show
End Sub

Private Sub mnuFees_Click()
    frmFees.Show
End Sub

Private Sub mnuFeesAndSchedules_Click()
    frmAssessSchedule.Show
End Sub

Private Sub mnuLedger_Click()
    frmLedger.Show
End Sub

Private Sub mnuPopulation_Click()
    frmPopulation.Show
End Sub

Private Sub mnuRegister_Click()
    frmRegister.Show
End Sub

Private Sub mnuSchedulingSystem_Click()
    frmSchedules.Show
End Sub

Private Sub mnuSchoolInfo_Click()
    frmSchoolInfo.Show
End Sub

Private Sub mnuSchoolYear_Click()
    frmSy.Show
End Sub

Private Sub mnuSemester_Click()
    frmSemester.Show
End Sub

Private Sub mnuTileHorizontal_Click()
    MDIForm1.Arrange vbTileHorizontal
End Sub

Private Sub mnuTileVertical_Click()
    MDIForm1.Arrange vbTileVertical
End Sub

Private Sub mnuStudentList_Click()
    frmStudentList.Show
End Sub

Private Sub mnuUsers_Click()
    frmUsers.Show
End Sub

Private Sub picSearch_Click()
    frmSearchStudent.Show
End Sub

'-- images switch
Private Sub picBox1_Click()
    'picHeader.Picture = LoadPicture(App.Path & "\images\header1_0.jpg")
    'picFooter.Picture = LoadPicture(App.Path & "\images\footer1_0.jpg")
    '-or-
    picHeader.Picture = imgHeader.ListImages(1).Picture
    picFooter.Picture = imgFooter.ListImages(1).Picture
    frmRegister.Show
End Sub
Private Sub picBox2_Click()
    'picHeader.Picture = LoadPicture(App.Path & "\images\header2_0.jpg")
    'picFooter.Picture = LoadPicture(App.Path & "\images\footer2_0.jpg")
    '-or-
    If LCase(User.UserType) <> "registrar" Then
        picHeader.Picture = imgHeader.ListImages(2).Picture
        picFooter.Picture = imgFooter.ListImages(2).Picture
        frmAssess.Show
    End If
End Sub
Private Sub picBox3_Click()
    'picHeader.Picture = LoadPicture(App.Path & "\images\header3_0.jpg")
    'picFooter.Picture = LoadPicture(App.Path & "\images\footer3_0.jpg")
    '-or-
    If LCase(User.UserType) <> "registrar" Then
        picHeader.Picture = imgHeader.ListImages(3).Picture
        picFooter.Picture = imgFooter.ListImages(3).Picture
        frmEnroll.Show
    End If
End Sub



Private Sub picBox1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'picBox1.Picture = LoadPicture(App.Path & "\images\b1_1.jpg")
    picBox1.Picture = imgBox1.ListImages(1).Picture
End Sub

Private Sub picBox2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'picBox2.Picture = LoadPicture(App.Path & "\images\b2_1.jpg")
    
    If LCase(User.UserType) <> "registrar" Then
        picBox2.Picture = imgBox1.ListImages(2).Picture
    Else
        picBox2.Picture = imgBoxX.ListImages(2).Picture
    End If
End Sub

Private Sub picBox3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'picBox3.Picture = LoadPicture(App.Path & "\images\b3_1.jpg")
    If LCase(User.UserType) <> "registrar" Then
        picBox3.Picture = imgBox1.ListImages(3).Picture
    Else
        picBox3.Picture = imgBoxX.ListImages(3).Picture
    End If
End Sub


Private Sub picSearch_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'picSearch.Picture = LoadPicture(App.Path & "\images\search_1.jpg")
    picSearch.Picture = imgSearch1.ListImages(1).Picture
End Sub

Private Sub picWall_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    picWall_Move
End Sub
Private Sub picWall_Move()
    'picBox1.Picture = LoadPicture(App.Path & "\images\b1_0.jpg")
    'picBox2.Picture = LoadPicture(App.Path & "\images\b2_0.jpg")
    'picBox3.Picture = LoadPicture(App.Path & "\images\b3_0.jpg")
    '--or
    picBox1.Picture = imgBox0.ListImages(1).Picture
    If LCase(User.UserType) <> "registrar" Then
        picBox2.Picture = imgBox0.ListImages(2).Picture
        picBox3.Picture = imgBox0.ListImages(3).Picture
    Else
        picBox2.Picture = imgBoxX.ListImages(2).Picture
        picBox3.Picture = imgBoxX.ListImages(3).Picture
    End If
    picSearch.Picture = imgSearch0.ListImages(1).Picture
End Sub
'-- eo: image switch




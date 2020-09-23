VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAssessSchedule 
   Caption         =   "Assessment and Schedules "
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmAssessSchedule.frx":0000
   ScaleHeight     =   2880
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtStatus 
      Appearance      =   0  'Flat
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
      Left            =   60
      MaxLength       =   20
      TabIndex        =   44
      Top             =   1080
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.OptionButton optCashBasis 
      Caption         =   "Cash Basis"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   43
      Top             =   1125
      Value           =   -1  'True
      Width           =   1680
   End
   Begin VB.OptionButton optInstallmentBasis 
      Caption         =   "Installment Basis"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   42
      Top             =   1440
      Width           =   1770
   End
   Begin VB.TextBox txtTotalCashBasis 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   7785
      MaxLength       =   20
      TabIndex        =   41
      Top             =   990
      Width           =   1410
   End
   Begin VB.TextBox txtTotalInstallmentBasis 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   7785
      MaxLength       =   20
      TabIndex        =   40
      Top             =   1350
      Width           =   1410
   End
   Begin VB.TextBox txtDownpayment 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   300
      Left            =   7785
      MaxLength       =   20
      TabIndex        =   39
      Top             =   1845
      Width           =   1410
   End
   Begin VB.TextBox txtEntrance 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "Entrance"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "rsFees"
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
      Height          =   285
      Left            =   7920
      MaxLength       =   20
      TabIndex        =   38
      Top             =   360
      Width           =   1590
   End
   Begin VB.TextBox txtTuitionFee 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "TuitionFee"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "rsFees"
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
      Height          =   285
      Left            =   7920
      MaxLength       =   20
      TabIndex        =   37
      Top             =   675
      Width           =   1590
   End
   Begin VB.TextBox txtRegistration 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "Registration"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "rsFees"
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
      Height          =   285
      Left            =   9585
      MaxLength       =   20
      TabIndex        =   36
      Top             =   450
      Width           =   1860
   End
   Begin VB.TextBox txtLibrary 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "Library"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "rsFees"
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
      Height          =   285
      Left            =   9585
      MaxLength       =   20
      TabIndex        =   35
      Top             =   765
      Width           =   1860
   End
   Begin VB.TextBox txtLaboratory 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "Laboratory"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "rsFees"
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
      Height          =   285
      Left            =   9585
      MaxLength       =   20
      TabIndex        =   34
      Top             =   1080
      Width           =   1860
   End
   Begin VB.TextBox txtAthleticFee 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "AthleticFee"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "rsFees"
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
      Height          =   285
      Left            =   9585
      MaxLength       =   20
      TabIndex        =   33
      Top             =   1395
      Width           =   1860
   End
   Begin VB.TextBox txtGuidanceAndCounselor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "GuidanceAndCounselor"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "rsFees"
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
      Height          =   285
      Left            =   9585
      MaxLength       =   20
      TabIndex        =   32
      Top             =   1710
      Width           =   1860
   End
   Begin VB.TextBox txtAffiliation 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "Affiliation"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "rsFees"
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
      Height          =   285
      Left            =   9585
      MaxLength       =   20
      TabIndex        =   31
      Top             =   2115
      Width           =   1860
   End
   Begin VB.TextBox txtNursingAudit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "NursingAudit"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "rsFees"
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
      Height          =   285
      Left            =   9585
      MaxLength       =   20
      TabIndex        =   30
      Top             =   2430
      Width           =   1860
   End
   Begin VB.TextBox txtMarineLaboratory 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "MarineLaboratory"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "rsFees"
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
      Height          =   285
      Left            =   9585
      MaxLength       =   20
      TabIndex        =   29
      Top             =   2745
      Width           =   1860
   End
   Begin VB.TextBox txtSpeechLab 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "SpeechLab"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "rsFees"
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
      Height          =   285
      Left            =   9585
      MaxLength       =   20
      TabIndex        =   28
      Top             =   3060
      Width           =   1860
   End
   Begin VB.TextBox txtHrmLab 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "HrmLab"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "rsFees"
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
      Height          =   285
      Left            =   9585
      MaxLength       =   20
      TabIndex        =   27
      Top             =   3375
      Width           =   1860
   End
   Begin VB.TextBox txtOjt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "Ojt"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "rsFees"
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
      Height          =   285
      Left            =   9585
      MaxLength       =   20
      TabIndex        =   26
      Top             =   3690
      Width           =   1860
   End
   Begin VB.TextBox txtRta 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "Rta"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "rsFees"
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
      Height          =   285
      Left            =   9585
      MaxLength       =   20
      TabIndex        =   25
      Top             =   4005
      Width           =   1860
   End
   Begin VB.TextBox txtHOn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "HOn"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "rsFees"
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
      Height          =   285
      Left            =   9585
      MaxLength       =   20
      TabIndex        =   24
      Top             =   4320
      Width           =   1860
   End
   Begin VB.TextBox txtMta 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "Mta"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "rsFees"
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
      Height          =   285
      Left            =   9585
      MaxLength       =   20
      TabIndex        =   23
      Top             =   4635
      Width           =   1860
   End
   Begin VB.TextBox txtIdNamePlate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "IdNamePlate"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "rsFees"
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
      Height          =   285
      Left            =   9585
      MaxLength       =   20
      TabIndex        =   22
      Top             =   4950
      Width           =   1860
   End
   Begin VB.TextBox txtPowerFee 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "PowerFee"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "rsFees"
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
      Height          =   285
      Left            =   9585
      MaxLength       =   20
      TabIndex        =   21
      Top             =   5580
      Width           =   1860
   End
   Begin VB.TextBox txtSdf 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "Sdf"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "rsFees"
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
      Height          =   285
      Left            =   9585
      MaxLength       =   20
      TabIndex        =   20
      Top             =   5265
      Width           =   1860
   End
   Begin VB.TextBox txtInternet 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "Internet"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "rsFees"
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
      Height          =   285
      Left            =   9585
      MaxLength       =   20
      TabIndex        =   19
      Top             =   5895
      Width           =   1860
   End
   Begin VB.TextBox txtInternship 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "Internship"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "rsFees"
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
      Height          =   285
      Left            =   9585
      MaxLength       =   20
      TabIndex        =   18
      Top             =   6210
      Width           =   1860
   End
   Begin VB.TextBox txtWaiver 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "Waiver"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "rsFees"
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
      Height          =   285
      Left            =   9585
      MaxLength       =   20
      TabIndex        =   17
      Top             =   6525
      Width           =   1860
   End
   Begin VB.TextBox txtNstp 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "Nstp"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "rsFees"
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
      Height          =   285
      Left            =   7920
      MaxLength       =   20
      TabIndex        =   16
      Top             =   6525
      Width           =   1635
   End
   Begin VB.TextBox txtRle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "Rle"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "rsFees"
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
      Height          =   285
      Left            =   7965
      MaxLength       =   20
      TabIndex        =   15
      Top             =   2115
      Width           =   1455
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
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
      Left            =   1845
      MaxLength       =   20
      TabIndex        =   14
      Top             =   1035
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.TextBox txtCourseCode 
      Appearance      =   0  'Flat
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
      Left            =   45
      MaxLength       =   20
      TabIndex        =   13
      Top             =   1890
      Visible         =   0   'False
      Width           =   1725
   End
   Begin MSMask.MaskEdBox meAssessNo 
      Height          =   465
      Left            =   4140
      TabIndex        =   12
      Top             =   1890
      Visible         =   0   'False
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   820
      _Version        =   393216
      Appearance      =   0
      BackColor       =   12640511
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtTotalLecUnits 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
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
      Left            =   1890
      MaxLength       =   20
      TabIndex        =   11
      Top             =   1980
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtTotalLabUnits 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
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
      Left            =   2475
      MaxLength       =   20
      TabIndex        =   10
      Top             =   1980
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox txtYearLevelId 
      Appearance      =   0  'Flat
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
      Left            =   4320
      MaxLength       =   20
      TabIndex        =   9
      Top             =   990
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton cmdSearch 
      Height          =   405
      Left            =   4095
      Picture         =   "frmAssessSchedule.frx":10380
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Search"
      Top             =   1440
      Width           =   450
   End
   Begin MSMask.MaskEdBox meStudentId 
      Height          =   420
      Left            =   1890
      TabIndex        =   0
      Top             =   1440
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   741
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
      ScaleWidth      =   5910
      TabIndex        =   5
      Top             =   2385
      Width           =   5970
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "Close"
         Height          =   315
         Left            =   4725
         TabIndex        =   3
         Top             =   60
         Width           =   1035
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "Generate"
         Height          =   315
         Left            =   3600
         TabIndex        =   2
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
         TabIndex        =   6
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
      ScaleWidth      =   5970
      TabIndex        =   1
      Top             =   0
      Width           =   5970
      Begin VB.Image Image2 
         Height          =   375
         Left            =   5580
         Picture         =   "frmAssessSchedule.frx":10755
         Top             =   0
         Width           =   390
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fees && Sched"
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
         Left            =   990
         TabIndex        =   4
         Top             =   120
         Width           =   4485
      End
      Begin VB.Image Image1 
         Height          =   780
         Left            =   0
         Picture         =   "frmAssessSchedule.frx":10C05
         Top             =   60
         Width           =   915
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Student #:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   630
      TabIndex        =   8
      Top             =   1530
      Width           =   1275
   End
End
Attribute VB_Name = "frmAssessSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsData As New ADODB.Recordset
Dim rsEnrolled As New ADODB.Recordset
Dim rsSchedules As New ADODB.Recordset
Dim rsFees As New ADODB.Recordset


Private Sub cmdGenerate_Click()
    ReportInformation.Sy = SchoolInformation.Sy
    ReportInformation.YearLevel = txtYearLevelId.Text 'dcYearLevel.Text
    ReportInformation.Semester = SchoolInformation.Semester
    
    ComputeFees txtCourseCode.Text, _
                txtYearLevelId.Text, _
                meAssessNo, _
                meStudentId, _
                txtName.Text, _
                txtStatus.Text, _
                optCashBasis.Value, _
                txtTotalLecUnits.Text, _
                txtTotalLabUnits.Text
                
    txtTotalInstallmentBasis.Text = Format(FeesBreakdown.TotalInstallmentBasis, "###,###.00") 'Imaginary Percentage
    txtDownpayment.Text = Format(FeesBreakdown.DownPayment, "###,###.00")  'Imaginary Percentage
    txtTotalCashBasis.Text = Format(FeesBreakdown.TotalCashBasis, "###,###.00")
    If rsEnrolled.State = adStateOpen Then
        If rsEnrolled.RecordCount > 0 Then
'            If LCase(SchoolInformation.Semester) = "summer" Then
'                Set rptAssessmentS.DataSource = rsEnrolled
'                rptAssessmentS.Sections("secFooter").Controls.Item("lblTotalLecUnits").Caption = Format(txtTotalLecUnits, "#0.0")
'                rptAssessmentS.Sections("secFooter").Controls.Item("lblTotalLabUnits").Caption = Format(txtTotalLabUnits, "#0.0")
'                rptAssessmentS.Show
'            Else
                Set rptAssessment1.DataSource = rsEnrolled
                rptAssessment1.Sections("secFooter").Controls.Item("lblTotalLecUnits").Caption = Format(txtTotalLecUnits, "#0.0")
                rptAssessment1.Sections("secFooter").Controls.Item("lblTotalLabUnits").Caption = Format(txtTotalLabUnits, "#0.0")
                rptAssessment1.Show
            'End If
        Else
            MsgBox "No report extracted.", vbInformation
        End If
    Else
        MsgBox "No report extracted.", vbInformation
    End If

    'Set rsEnrolled = Nothing
End Sub

Private Sub cmdSearch_Click()
    frmSearchStudent.txtCallingForm = "frmAssessSchedule"
    frmSearchStudent.Show vbModal
End Sub

Public Sub meStudentId_LostFocus()
    strSql = "SELECT * " & _
             "FROM dbo.Students AS s " & _
                  "JOIN dbo.Courses AS c ON (s.CourseCode = c.CourseCode) " & _
                  "JOIN dbo.YearLevel AS yl ON (s.YearLevelId = yl.YearLevelId) " & _
                  "JOIN dbo.Semester AS sem ON (s.SemesterId = sem.SemesterId) " & _
            "WHERE s.StudentId = '" & meStudentId.Text & "';"

    
    
    Set rsData = GetRecordset(strSql)
    If rsData.State = adStateOpen Then
        If rsData.RecordCount > 0 Then
            txtName.Text = rsData!Lastname & ", " & rsData!Firstname & " " & rsData!Middlename
            txtYearLevelId.Text = rsData!YearLevelId
            txtCourseCode.Text = rsData!CourseCode
            txtStatus.Text = rsData!Status
        End If
    End If
    Set rsData = Nothing
    
    strSql = "SELECT SUM(LecUnits) AS Lec, SUM(LabUnits) AS Lab, CashBasis " & _
             "FROM dbo.Enrolled en " & _
             "INNER JOIN dbo.Schedules sc ON (en.SchedCode = sc.SchedCode) " & _
             "INNER JOIN dbo.Subjects sj ON (sc.SubjectCode = sj.SubjectCode) " & _
             "WHERE AssessNo = " & meAssessNo.Text & _
                    " AND (en.Status = 'P' OR en.Status = 'C')" & _
             "GROUP BY CashBasis;"
    Set rsData = GetRecordset(strSql)
    If rsData.State = adStateOpen Then
        If rsData.RecordCount > 0 Then
            txtTotalLecUnits.Text = rsData!Lec
            txtTotalLabUnits.Text = rsData!Lab
            optCashBasis.Value = rsData!CashBasis
            optInstallmentBasis.Value = (rsData!CashBasis = 0)
        End If
    End If
    Set rsData = Nothing
    
    'EnrollId, SchedCode, Subject, Lec, Lab, Time, Days, Room,  Instructor, Section
    strSql = "SELECT EnrollId, sj.SubjectCode  AS Code, SubjectDesc AS [Subject Description], LecUnits AS Lec, LabUnits AS Lab, (sc.TimeSchedStart + '-' + sc.TimeSchedEnd) AS [Time], DaysSched AS Days, RoomNo AS Room, Instructor, sc.SchedCode, en.StudentId " & _
             "FROM dbo.Enrolled en " & _
             "INNER JOIN dbo.Schedules sc ON (en.SchedCode = sc.SchedCode) " & _
             "INNER JOIN dbo.Subjects sj ON (sc.SubjectCode = sj.SubjectCode) " & _
             "INNER JOIN dbo.Rooms rm ON (sc.RoomId = rm.RoomId) " & _
             "INNER JOIN dbo.Instructors tr ON (sc.InstructorId = tr.InstructorId) " & _
             "WHERE AssessNo = " & meAssessNo & " AND (en.Status = 'P' OR en.Status = 'C') ORDER BY Days;" 'P=Pending, C-Confirmed
             
    'Debug.Print strSql
    Set rsEnrolled = GetRecordset(strSql)
    'Set rsEnrolled = Nothing
End Sub



'Icon on Message vbCritical 16, vbQuestion 32, vbExclamation 48, vbInformation 64
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set rsData = Nothing
    Set rsEnrolled = Nothing
    Set rsSchedules = Nothing
    Set rsFees = Nothing
End Sub

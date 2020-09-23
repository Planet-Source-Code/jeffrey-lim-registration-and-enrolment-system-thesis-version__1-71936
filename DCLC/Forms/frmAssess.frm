VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAssess 
   Caption         =   "Assessment"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmAssess.frx":0000
   ScaleHeight     =   533
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   759
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtStatus 
      Appearance      =   0  'Flat
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
      Left            =   990
      MaxLength       =   20
      TabIndex        =   75
      Top             =   2070
      Width           =   1860
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
      Left            =   11475
      MaxLength       =   20
      TabIndex        =   72
      Top             =   2700
      Width           =   1455
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
      Left            =   11430
      MaxLength       =   20
      TabIndex        =   71
      Top             =   7110
      Width           =   1635
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
      Left            =   13095
      MaxLength       =   20
      TabIndex        =   70
      Top             =   7110
      Width           =   1860
   End
   Begin VB.PictureBox Picture1 
      Height          =   1320
      Left            =   7920
      ScaleHeight     =   1260
      ScaleWidth      =   3330
      TabIndex        =   63
      Top             =   945
      Width           =   3390
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
         Left            =   90
         TabIndex        =   74
         Top             =   540
         Width           =   1770
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
         Left            =   90
         TabIndex        =   73
         Top             =   90
         Value           =   -1  'True
         Width           =   1680
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
         Left            =   1890
         MaxLength       =   20
         TabIndex        =   66
         Top             =   945
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
         Left            =   1890
         MaxLength       =   20
         TabIndex        =   65
         Top             =   585
         Width           =   1410
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
         Left            =   1890
         MaxLength       =   20
         TabIndex        =   64
         Top             =   90
         Width           =   1410
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         X1              =   90
         X2              =   3285
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Downpayment"
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
         Left            =   360
         TabIndex        =   69
         Top             =   945
         Width           =   1410
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
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
         Height          =   285
         Left            =   270
         TabIndex        =   68
         Top             =   4035
         Width           =   1455
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
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
         Height          =   285
         Left            =   270
         TabIndex        =   67
         Top             =   3675
         Width           =   1005
      End
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
      Left            =   13095
      MaxLength       =   20
      TabIndex        =   62
      Top             =   6795
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
      Left            =   13095
      MaxLength       =   20
      TabIndex        =   61
      Top             =   6480
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
      Left            =   13095
      MaxLength       =   20
      TabIndex        =   60
      Top             =   5850
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
      Left            =   13095
      MaxLength       =   20
      TabIndex        =   59
      Top             =   6165
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
      Left            =   13095
      MaxLength       =   20
      TabIndex        =   58
      Top             =   5535
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
      Left            =   13095
      MaxLength       =   20
      TabIndex        =   57
      Top             =   5220
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
      Left            =   13095
      MaxLength       =   20
      TabIndex        =   56
      Top             =   4905
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
      Left            =   13095
      MaxLength       =   20
      TabIndex        =   55
      Top             =   4590
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
      Left            =   13095
      MaxLength       =   20
      TabIndex        =   54
      Top             =   4275
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
      Left            =   13095
      MaxLength       =   20
      TabIndex        =   53
      Top             =   3960
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
      Left            =   13095
      MaxLength       =   20
      TabIndex        =   52
      Top             =   3645
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
      Left            =   13095
      MaxLength       =   20
      TabIndex        =   51
      Top             =   3330
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
      Left            =   13095
      MaxLength       =   20
      TabIndex        =   50
      Top             =   3015
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
      Left            =   13095
      MaxLength       =   20
      TabIndex        =   49
      Top             =   2700
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
      Left            =   13095
      MaxLength       =   20
      TabIndex        =   48
      Top             =   2295
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
      Left            =   13095
      MaxLength       =   20
      TabIndex        =   47
      Top             =   1980
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
      Left            =   13095
      MaxLength       =   20
      TabIndex        =   46
      Top             =   1665
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
      Left            =   13095
      MaxLength       =   20
      TabIndex        =   45
      Top             =   1350
      Width           =   1860
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
      Left            =   13095
      MaxLength       =   20
      TabIndex        =   44
      Top             =   1035
      Width           =   1860
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
      Left            =   11430
      MaxLength       =   20
      TabIndex        =   43
      Top             =   1260
      Width           =   1590
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
      Left            =   11430
      MaxLength       =   20
      TabIndex        =   42
      Top             =   945
      Width           =   1590
   End
   Begin VB.TextBox txtInstructor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataField       =   "Instructor"
      DataSource      =   "rsSchedules"
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
      Left            =   8685
      MaxLength       =   20
      TabIndex        =   41
      Top             =   2835
      Width           =   2130
   End
   Begin VB.TextBox txtRoomNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataField       =   "RoomNo"
      DataSource      =   "rsSchedules"
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
      Left            =   4590
      MaxLength       =   20
      TabIndex        =   40
      Top             =   2835
      Width           =   735
   End
   Begin VB.TextBox txtEnrollId 
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
      Left            =   2880
      MaxLength       =   20
      TabIndex        =   38
      Top             =   1710
      Visible         =   0   'False
      Width           =   1860
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
      Left            =   5940
      MaxLength       =   20
      TabIndex        =   37
      Top             =   7155
      Width           =   510
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
      Left            =   5355
      MaxLength       =   20
      TabIndex        =   36
      Top             =   7155
      Width           =   555
   End
   Begin VB.CommandButton cmdRemove 
      Enabled         =   0   'False
      Height          =   330
      Left            =   10845
      Picture         =   "frmAssess.frx":10380
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Remove from Temp"
      Top             =   2520
      Width           =   465
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   330
      Left            =   10845
      Picture         =   "frmAssess.frx":1081E
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Add to Temp"
      Top             =   2835
      Width           =   465
   End
   Begin VB.TextBox txtDays 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataField       =   "DaysSched"
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
      Left            =   7830
      MaxLength       =   20
      TabIndex        =   30
      Top             =   2835
      Width           =   825
   End
   Begin VB.TextBox txtTimeSched 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   6525
      MaxLength       =   20
      TabIndex        =   29
      Top             =   2835
      Width           =   1275
   End
   Begin VB.TextBox txtLabUnits 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataField       =   "LabUnits"
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
      Height          =   315
      Left            =   5940
      MaxLength       =   20
      TabIndex        =   28
      Top             =   2835
      Width           =   555
   End
   Begin VB.TextBox txtLecUnits 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataField       =   "LecUnits"
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
      Height          =   315
      Left            =   5355
      MaxLength       =   20
      TabIndex        =   27
      Top             =   2835
      Width           =   555
   End
   Begin MSDataListLib.DataCombo dcSubject 
      Height          =   315
      Left            =   45
      TabIndex        =   4
      Top             =   2835
      Width           =   4530
      _ExtentX        =   7990
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
      Left            =   990
      TabIndex        =   3
      Top             =   1710
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
      Left            =   990
      TabIndex        =   2
      Top             =   1350
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin MSMask.MaskEdBox meStudentId 
      Height          =   330
      Left            =   990
      TabIndex        =   1
      Top             =   990
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
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      DataField       =   "Name"
      DataSource      =   "adoStudents"
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
      Left            =   2610
      MaxLength       =   35
      TabIndex        =   16
      Top             =   990
      Width           =   4905
   End
   Begin VB.CommandButton cmdSearch 
      Height          =   315
      Left            =   7515
      Picture         =   "frmAssess.frx":10CCD
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Search"
      Top             =   990
      Width           =   360
   End
   Begin VB.PictureBox picFooter 
      Align           =   2  'Align Bottom
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   11325
      TabIndex        =   11
      Top             =   7500
      Width           =   11385
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "Close"
         Height          =   315
         Left            =   10170
         TabIndex        =   8
         Top             =   60
         Width           =   1035
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Commit"
         Height          =   315
         Left            =   9090
         TabIndex        =   7
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
         TabIndex        =   12
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
      TabIndex        =   9
      Top             =   0
      Width           =   11385
      Begin MSMask.MaskEdBox meAssessNo 
         Height          =   465
         Left            =   9270
         TabIndex        =   0
         Top             =   180
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
         Caption         =   "Assess#:"
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
         Left            =   7875
         TabIndex        =   15
         Top             =   225
         Width           =   1365
      End
      Begin VB.Image Image2 
         Height          =   375
         Left            =   11025
         Picture         =   "frmAssess.frx":110A2
         Top             =   0
         Width           =   390
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Assessment"
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
         Width           =   3630
      End
      Begin VB.Image Image1 
         Height          =   780
         Left            =   0
         Picture         =   "frmAssess.frx":11552
         Top             =   60
         Width           =   915
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAssess.frx":11D43
      Height          =   3870
      Left            =   45
      TabIndex        =   18
      Top             =   3240
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   6826
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
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Type:"
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
      Left            =   180
      TabIndex        =   76
      Top             =   2070
      Width           =   825
   End
   Begin VB.Image imgPreview 
      Height          =   360
      Left            =   7515
      Picture         =   "frmAssess.frx":11D5D
      Top             =   1845
      Width           =   360
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Room"
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
      Left            =   4590
      TabIndex        =   39
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "units"
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
      Left            =   4545
      TabIndex        =   35
      Top             =   7110
      Width           =   735
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "system"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   1530
      TabIndex        =   34
      Top             =   0
      Width           =   1320
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "External"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   465
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   1545
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Registered"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   465
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   2040
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "location"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   2025
      TabIndex        =   31
      Top             =   0
      Width           =   1995
   End
   Begin VB.Label Label27 
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
      Height          =   285
      Left            =   180
      TabIndex        =   26
      Top             =   1710
      Width           =   690
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Lab"
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
      Left            =   5940
      TabIndex        =   25
      Top             =   2520
      Width           =   555
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "INSTRUCTOR"
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
      Left            =   8685
      TabIndex        =   24
      Top             =   2520
      Width           =   2130
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Left            =   3645
      TabIndex        =   23
      Top             =   7110
      Width           =   870
   End
   Begin VB.Label Label7 
      BackColor       =   &H000080FF&
      Caption         =   " Subject"
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
      Left            =   45
      TabIndex        =   22
      Top             =   2520
      Width           =   4515
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Time"
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
      Left            =   6525
      TabIndex        =   21
      Top             =   2520
      Width           =   1275
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Days"
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
      Left            =   7830
      TabIndex        =   20
      Top             =   2520
      Width           =   825
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   " Lec"
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
      Left            =   5355
      TabIndex        =   19
      Top             =   2520
      Width           =   555
   End
   Begin VB.Label Label2 
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
      Height          =   330
      Left            =   165
      TabIndex        =   17
      Top             =   1350
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Student :"
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
      Left            =   165
      TabIndex        =   14
      Top             =   1035
      Width           =   825
   End
End
Attribute VB_Name = "frmAssess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsData As New ADODB.Recordset
Dim rsEnrolled As New ADODB.Recordset
Dim rsSchedules As New ADODB.Recordset
Dim rsFees As New ADODB.Recordset

Private Sub cmdRemove_Click()
    If txtEnrollId.Text <> "" Then
        strSql = "DELETE FROM dbo.Enrolled WHERE (Status = 'P' OR Status 'C') AND EnrollId = " & txtEnrollId.Text & ";"
        RunSql (strSql)
        SetSubjectsAssessed
        txtEnrollId.Text = ""
    End If
End Sub




Private Sub dcSubject_Change()
    On Error Resume Next
    strSql = "SELECT LecUnits, LabUnits, (sc.TimeSchedStart + '-' + sc.TimeSchedEnd) AS TimeSched, DaysSched, RoomNo, Instructor " & _
             "FROM dbo.Schedules sc " & _
             "INNER JOIN dbo.Subjects sj ON (sc.SubjectCode = sj.SubjectCode) " & _
             "INNER JOIN dbo.Rooms rm ON (sc.RoomId = rm.RoomId) " & _
             "INNER JOIN dbo.Instructors tr ON (sc.InstructorId = tr.InstructorId) " & _
             "WHERE SchedCode = " & dcSubject.BoundText & ";"
             
    'Debug.Print strSql
    Set rsSchedules = GetRecordset(strSql)
  
     If rsSchedules.State = adStateOpen Then
        If rsSchedules.RecordCount > 0 Then
            Set txtRoomNo.DataSource = rsSchedules
            txtTimeSched.Text = rsSchedules!TimeSched 'work around
            'txtTimeSched.DataField = "TimeSched" 'w/ bug
            'Set txtTimeSched.DataSource = rsSchedules
            Set txtLecUnits.DataSource = rsSchedules
            Set txtLabUnits.DataSource = rsSchedules
            Set txtDays.DataSource = rsSchedules
            Set txtInstructor.DataSource = rsSchedules
        End If
    End If
    
    'Check if subject is already in dbo.Enrolled table
    Set rsData = GetRecordset("SELECT EnrollId FROM dbo.Enrolled " & _
                              "WHERE SchedCode = " & dcSubject.BoundText & " AND " & _
                                    "AssessNo = " & meAssessNo & " AND " & _
                                    "StudentId = '" & meStudentId & "';")
    txtEnrollId.Text = ""
    cmdRemove.Enabled = False
    If rsData.State = adStateOpen Then
        If rsData.RecordCount > 0 Then
            txtEnrollId.Text = rsData!EnrollId
            cmdRemove.Enabled = True
        End If
    End If
    Set rsData = Nothing
  'Set DataGrid1.DataSource = rsEnrolled
  
End Sub

Private Sub cmdAdd_Click()
    If dcSubject.Text <> "" And EntriesValid And txtEnrollId.Text = "" Then
        strSql = "Declare @FeesId1 integer; "
        strSql = strSql + "Set @FeesId1 = (SELECT FeesId " & _
                                          "FROM dbo.Fees " & _
                                          "WHERE CourseCode = '" & dcCourses.BoundText & "' AND " & _
                                                "CurrentYearLevelId = " & dcYearLevel.BoundText & " AND " & _
                                                "CurrentSyId = " & SchoolInformation.CurrentSyId & " AND " & _
                                                "CurrentSemesterId = " & SchoolInformation.CurrentSemesterId & ");"
                                                
        strSql = strSql + "INSERT INTO dbo.Enrolled(AssessNo, StudentId, SchedCode, FeesId, CashBasis) " & _
                          "VALUES(" & meAssessNo & ", '" & meStudentId & "', '" & dcSubject.BoundText & "', @FeesId1, '" & IIf(optCashBasis.Value, 1, 0) & "');"
                          
        'Debug.Print strSql
        
'        If txtEnrollId.Text <> "" Then
'            strSql = "UPDATE dbo.Enrolled SET  Status = 'Deleted' " & _
'                        "WHERE StudentId = '" & meStudentId.Text & "'; "
'        End If
        
        'Execute SQL Command
        RunSql (strSql)
        
        SetSubjectsAssessed

        meStudentId.Enabled = False
        cmdSearch.Enabled = False
        dcSubject.SetFocus
    End If
End Sub

Private Sub GetFees()
    strSql = "SELECT * FROM dbo.Fees " & _
             "WHERE CourseCode = '" & dcCourses.BoundText & "' AND " & _
                   "CurrentSyId = " & SchoolInformation.CurrentSyId & " AND " & _
                   "CurrentSemesterId = " & SchoolInformation.CurrentSemesterId & " AND " & _
                   "CurrentYearLevelId = " & dcYearLevel.BoundText & " ;"
     
    Set rsFees = GetRecordset(strSql)
    If rsFees.State = adStateOpen Then
        If rsFees.RecordCount > 0 Then
            '-- For Debugging purpose only
            Set txtEntrance.DataSource = rsFees
            Set txtTuitionFee.DataSource = rsFees
            
            Set txtRegistration.DataSource = rsFees
            Set txtLibrary.DataSource = rsFees
            Set txtLaboratory.DataSource = rsFees
            Set txtAthleticFee.DataSource = rsFees
            Set txtGuidanceAndCounselor.DataSource = rsFees
            
            Set txtRle.DataSource = rsFees
            Set txtAffiliation.DataSource = rsFees
            Set txtNursingAudit.DataSource = rsFees
            Set txtMarineLaboratory.DataSource = rsFees
            Set txtSpeechLab.DataSource = rsFees
            Set txtHrmLab.DataSource = rsFees
            Set txtOjt.DataSource = rsFees
            Set txtRta.DataSource = rsFees
            Set txtHOn.DataSource = rsFees
            Set txtMta.DataSource = rsFees
            Set txtIdNamePlate.DataSource = rsFees
            Set txtSdf.DataSource = rsFees
            Set txtPowerFee.DataSource = rsFees
            Set txtInternet.DataSource = rsFees
            Set txtInternship.DataSource = rsFees
            Set txtWaiver.DataSource = rsFees
            Set txtNstp.DataSource = rsFees
            '-- eo: For Debugging purpose only
            
             ComputeFees dcCourses.BoundText, _
                        dcYearLevel.BoundText, _
                        meAssessNo, _
                        meStudentId, _
                        txtName, _
                        txtStatus, _
                        optCashBasis.Value, _
                        txtTotalLecUnits, _
                        txtTotalLabUnits
                      
            txtTotalInstallmentBasis.Text = Format(FeesBreakdown.TotalInstallmentBasis, "###,###.00") 'Imaginary Percentage
            txtDownpayment.Text = Format(FeesBreakdown.DownPayment, "###,###.00")  'Imaginary Percentage
            txtTotalCashBasis.Text = Format(FeesBreakdown.TotalCashBasis, "###,###.00")
        Else
            MsgBox "No fee schedule found", vbInformation
        End If
    End If
    Set rsFees = Nothing
End Sub

Private Sub SetSubjectsAssessed()
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
  
    If rsEnrolled.State = adStateOpen Then
        If rsEnrolled.RecordCount > 0 Then
            'Set txtTotalLecUnits.DataSource = rsEnrolled
            'Set txtTotalLabUnits.DataSource = rsEnrolled
            
            '-- redundant as workaround due to datagrid subscript out of range
            strSql = "SELECT SUM(LecUnits) AS Lec, SUM(LabUnits) AS Lab " & _
                     "FROM dbo.Enrolled en " & _
                     "INNER JOIN dbo.Schedules sc ON (en.SchedCode = sc.SchedCode) " & _
                     "INNER JOIN dbo.Subjects sj ON (sc.SubjectCode = sj.SubjectCode) " & _
                     "WHERE AssessNo = " & meAssessNo & " AND (en.Status = 'P' OR en.Status = 'C');" 'P=Pending, C-Confirmed

            Set rsData = GetRecordset(strSql)
            If rsData.State = adStateOpen Then
                If rsData.RecordCount > 0 Then
                    txtTotalLecUnits.Text = rsData!Lec
                    txtTotalLabUnits.Text = rsData!Lab
                End If
            End If
            Set rsData = Nothing
            '--
        Else
            'enable/disable controls here
            meStudentId.Enabled = True
            dcCourses.Enabled = True
            dcYearLevel.Enabled = True
            cmdRemove.Enabled = False
        End If
    End If
  
    Set DataGrid1.DataSource = rsEnrolled
    DataGrid1.Columns(2).Width = 200
    DataGrid1.Columns(3).Alignment = dbgCenter
    DataGrid1.Columns(4).Alignment = dbgCenter
    
    GetFees
End Sub


Private Sub cmdSave_Click()
    If EntriesValid Then
        If MsgBox("Do you want to " & cmdSave.Caption & " record?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
                strSql = "UPDATE dbo.Enrolled SET Status = 'C', " & _
                                                 "CashBasis = '" & IIf(optCashBasis.Value, 1, 0) & "', " & _
                                                 "StudentId = '" & meStudentId.Text & "' " & _
                         "WHERE AssessNo = " & meAssessNo & "; "
                strSql = strSql + "UPDATE dbo.Students SET CourseCode = '" & dcCourses.BoundText & "', " & _
                                                          "YearLevelId = " & dcYearLevel.BoundText & ", " & _
                                                          "SemesterId = " & SchoolInformation.CurrentSemesterId & ", " & _
                                                          "DateEnrolled = '" & Date & "' " & _
                                  "WHERE StudentId = '" & meStudentId & "'; "
                strSql = strSql + "DELETE FROM dbo.Ledger WHERE ReceiptNo = '" & meAssessNo & "*'; "
                                 '"DELETE FROM dbo.Ledger WHERE StudentId = '" & meStudentId & "' AND " & _
                                 '                              "SyId = " & SchoolInformation.CurrentSyId & " AND " & _
                                 '                              "SemesterId = " & SchoolInformation.CurrentSemesterId & "; "
                strSql = strSql + "INSERT INTO dbo.Ledger(ReceiptNo, StudentId, SyId, YearLevelId, SemesterId, Debit, Particular, TranDate, PostedBy) " & _
                                  "VALUES('" & meAssessNo & "*', '" & meStudentId & "', " & SchoolInformation.CurrentSyId & ", " & _
                                          dcYearLevel.BoundText & ", " & SchoolInformation.CurrentSemesterId & ", " & _
                                          IIf(optCashBasis, FeesBreakdown.TotalCashBasis, FeesBreakdown.TotalInstallmentBasis + FeesBreakdown.DownPayment) & ", " & _
                                        "'Enrollment-" & meAssessNo & "', '" & Date & "', '" & User.UserId & "'); "
                         
            RunSql (strSql) 'Execute SQL Command

            MsgBox "Please take note of Assessment # : " & meAssessNo, vbExclamation
            
            imgPreview_Click
            
            ClearEntries Me
            Call UpdatePK("AssessNo", cmdSave.Caption) 'Position this before ValidateAccessLevel.Insert
            ValidateAccessLevel Me, "Commit"
            
            meAssessNo_GotFocus
            cmdSearch.SetFocus
            'EmulateEnter 13
            
        End If
    End If
End Sub


Private Sub dcCourses_Change()
    BindDataCombo "SELECT *, sj.SubjectDesc + ' (' + CAST(sc.SchedCode AS VARCHAR) + ')' AS SubjectDesc2 FROM dbo.Schedules sc, dbo.Subjects sj " & _
                   "WHERE sc.SubjectCode = sj.SubjectCode AND " & _
                         "sc.SyId = " & SchoolInformation.CurrentSyId & " AND " & _
                         "sc.SemesterId = " & SchoolInformation.CurrentSemesterId & " AND " & _
                         "(sc.CourseCode = '" & dcCourses.BoundText & "' OR sc.CourseCode IS NULL) " & _
                         " ORDER BY sj.SubjectDesc; ", "SubjectDesc2", dcSubject, "SchedCode", False 'Bind Subjects
End Sub

Private Sub Form_Load()
    BindDataCombo "SELECT * FROM dbo.Courses", "CourseDesc", dcCourses, "CourseCode", True 'Bind Courses
    BindDataCombo "SELECT * FROM dbo.YearLevel", "YearLevel", dcYearLevel, "YearLevelId", True 'Bind School Year
'    BindDataCombo "SELECT *, sj.SubjectDesc + ' (' + CAST(sc.SchedCode AS VARCHAR) + ')' AS SubjectDesc2 FROM dbo.Schedules sc, dbo.Subjects sj " & _
'                   "WHERE sc.SubjectCode = sj.SubjectCode AND " & _
'                         "sc.SyId = " & SchoolInformation.CurrentSyId & " AND " & _
'                         "sc.SemesterId = " & SchoolInformation.CurrentSemesterId & " ORDER BY sj.SubjectDesc; ", "SubjectDesc2", dcSubject, "SchedCode", False 'Bind Subjects
'    BindDataCombo "SELECT * FROM dbo.Semester", "Semester", dcSemester, "SemesterId", True 'Bind Semester
    meAssessNo_GotFocus
    EmulateEnter 13
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    txtEnrollId.Text = ""
    cmdRemove.Enabled = False
    If rsEnrolled.State = adStateOpen Then
        If Not rsEnrolled.BOF And Not rsEnrolled.EOF Then
            txtEnrollId.Text = rsEnrolled!EnrollId
            dcSubject.BoundText = rsEnrolled!SchedCode
            cmdRemove.Enabled = True
        End If
    End If
End Sub



Private Sub imgPreview_Click()
    ReportInformation.Sy = SchoolInformation.Sy
    ReportInformation.YearLevel = dcYearLevel.Text
    ReportInformation.Semester = SchoolInformation.Semester
    GetFees
    If rsEnrolled.State = adStateOpen Then
        If rsEnrolled.RecordCount > 0 Then
'            If LCase(SchoolInformation.Semester) = "summer" Then
'                Set rptAssessmentS.DataSource = rsEnrolled
'                rptAssessmentS.Sections("secFooter").Controls.Item("lblTotalLecUnits").Caption = Format(txtTotalLecUnits, "#0.0")
'                rptAssessmentS.Sections("secFooter").Controls.Item("lblTotalLabUnits").Caption = Format(txtTotalLabUnits, "#0.0")
'                rptAssessmentS.Show vbModal
'            Else
                Set rptAssessment1.DataSource = rsEnrolled
                rptAssessment1.Sections("secFooter").Controls.Item("lblTotalLecUnits").Caption = Format(txtTotalLecUnits, "#0.0")
                rptAssessment1.Sections("secFooter").Controls.Item("lblTotalLabUnits").Caption = Format(txtTotalLabUnits, "#0.0")
                rptAssessment1.Show vbModal
'            End If
        Else
            MsgBox "No report extracted.", vbInformation
        End If
    Else
        MsgBox "No report extracted.", vbInformation
    End If

    'Set rsEnrolled = Nothing
End Sub

Private Sub meAssessNo_GotFocus()
    ClearEntries Me
    ValidateAccessLevel Me, "Commit"
    meStudentId.Enabled = True
    cmdSearch.Enabled = True
    meAssessNo.Text = GenerateNextPK("AssessNo")
    FocusMe (meAssessNo)
    Set DataGrid1.DataSource = Nothing
    DataGrid1.Refresh
    'SetSubjectsAssessed
End Sub

Private Sub meAssessNo_LostFocus()
    strSql = "SELECT en.StudentId, en.Status, en.CashBasis, s.DateEnrolled " & _
             "FROM dbo.Enrolled en " & _
             "INNER JOIN dbo.Students s ON (en.StudentId = s.StudentId) " & _
             "WHERE en.AssessNo = " & meAssessNo & ";"
     
    Set rsData = GetRecordset(strSql)
    If rsData.State = adStateOpen Then
        If rsData.RecordCount > 0 Then
            If rsData!Status <> "P" And rsData!DateEnrolled <> Date Then
                MsgBox "Assessment # was already been confirmed as enrolled on " & rsData!DateEnrolled & " . Pls assign different assessment #. ", vbInformation
            Else
                meStudentId.Mask = ""
                meStudentId.Text = ""
                meStudentId = rsData!StudentId
                optCashBasis.Value = rsData!CashBasis
                optInstallmentBasis.Value = (rsData!CashBasis = 0)
                meStudentId.Enabled = False
                cmdSearch.Enabled = False
                meStudentId_LostFocus
            End If
        End If
    End If
    Set rsData = Nothing
End Sub

Public Sub meStudentId_LostFocus()
    strSql = "SELECT *, (Lastname + ', ' + Firstname + ' ' + Middlename) AS Name, Status " & _
             "FROM dbo.Students AS s " & _
                  "JOIN dbo.Courses AS c ON (s.CourseCode = c.CourseCode) " & _
                  "JOIN dbo.YearLevel AS yl ON (s.YearLevelId = yl.YearLevelId) " & _
                  "JOIN dbo.Semester AS sem ON (s.SemesterId = sem.SemesterId) " & _
            "WHERE s.StudentId = '" & meStudentId.Text & "';"
     
    Set rsData = GetRecordset(strSql)
    If rsData.State = adStateOpen Then
        If rsData.RecordCount > 0 Then
            txtName.Text = rsData!Name
            txtStatus.Text = rsData!Status
            dcCourses.BoundText = rsData!CourseCode
            dcYearLevel.BoundText = rsData!YearLevelId
            SetSubjectsAssessed
            
            
        End If
    End If
    Set rsData = Nothing
End Sub


Private Sub cmdSearch_Click()
    frmSearchStudent.txtCallingForm = "frmAssess"
    frmSearchStudent.Show vbModal
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    MakeTransparent Me.hwnd, 255 'Unfade form
End Sub

Private Sub Form_Deactivate()
    MakeTransparent Me.hwnd, 200 'Fade Form
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsData = Nothing
    Set rsEnrolled = Nothing
    Set rsSchedules = Nothing
    Set rsFees = Nothing
End Sub

'Tab emulation
Private Sub meAssessNo_KeyPress(KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub

Private Sub meStudentId_GotFocus()
    FocusMe (meStudentId)
End Sub

Private Sub meStudentId_KeyPress(KeyAscii As Integer)
    EmulateEnter KeyAscii
End Sub

Private Sub dcSubject_KeyPress(KeyAscii As Integer)
    EmulateEnter KeyAscii
End Sub

Private Function EntriesValid() As Boolean
    EntriesValid = False
    If meAssessNo.Text = "" Then
        MsgBox ("Assessment # required!")
        meAssessNo.SetFocus
        Exit Function
    End If
    If meStudentId.Text = "__-_______" Then
        MsgBox ("Student # required!")
        meStudentId.SetFocus
        Exit Function
    End If
    EntriesValid = True
End Function

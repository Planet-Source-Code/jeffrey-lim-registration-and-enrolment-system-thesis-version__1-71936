VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmFees 
   Caption         =   "Fees Schedule"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmFees.frx":0000
   ScaleHeight     =   582.105
   ScaleMode       =   0  'User
   ScaleWidth      =   748
   StartUpPosition =   3  'Windows Default
   Begin MSDataListLib.DataCombo dcCourses 
      Height          =   315
      Left            =   1035
      TabIndex        =   0
      Top             =   990
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6405
      Left            =   90
      TabIndex        =   120
      Top             =   1350
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   11298
      _Version        =   393216
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
      TabCaption(0)   =   "Entrance, Tuition && Other Fees"
      TabPicture(0)   =   "frmFees.frx":10380
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label28"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label27"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label26"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label25"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label29"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label30"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label31"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label32"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label33"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label34"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label58"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label59"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label60"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label61"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label62"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label63"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "meRegistration(4)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "meRegistration(3)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "meRegistration(2)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "meRegistration(1)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "meEntrance(4)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "meEntrance(3)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "meEntrance(2)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "meEntrance(1)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "meGuidanceAndCounselor(4)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "meGuidanceAndCounselor(3)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "meGuidanceAndCounselor(2)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "meGuidanceAndCounselor(1)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "meAthleticFee(4)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "meAthleticFee(3)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "meAthleticFee(2)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "meAthleticFee(1)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "meLaboratory(4)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "meLaboratory(3)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "meLaboratory(2)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "meLaboratory(1)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "meGuidanceAndCounselor(0)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "meAthleticFee(0)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "meLaboratory(0)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "meLibrary(4)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "meLibrary(3)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "meLibrary(2)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "meLibrary(1)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "meLibrary(0)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "meRegistration(0)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "meTuitionFee(4)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "meTuitionFee(3)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "meTuitionFee(2)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "meTuitionFee(1)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "meTuitionFee(0)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "meEntrance(0)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txtOtherFeesTotal(0)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "txtOtherFeesTotal(1)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "txtOtherFeesTotal(2)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "txtOtherFeesTotal(3)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "txtOtherFeesTotal(4)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).ControlCount=   56
      TabCaption(1)   =   "Miscellaneous"
      TabPicture(1)   =   "frmFees.frx":1039C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label35"
      Tab(1).Control(1)=   "Label38"
      Tab(1).Control(2)=   "Label39"
      Tab(1).Control(3)=   "Label40"
      Tab(1).Control(4)=   "Label41"
      Tab(1).Control(5)=   "Label42"
      Tab(1).Control(6)=   "Label36"
      Tab(1).Control(7)=   "Label37"
      Tab(1).Control(8)=   "Label43"
      Tab(1).Control(9)=   "Label44"
      Tab(1).Control(10)=   "Label45"
      Tab(1).Control(11)=   "Label46"
      Tab(1).Control(12)=   "Label47"
      Tab(1).Control(13)=   "Label48"
      Tab(1).Control(14)=   "Label49"
      Tab(1).Control(15)=   "Label50"
      Tab(1).Control(16)=   "Label51"
      Tab(1).Control(17)=   "Label52"
      Tab(1).Control(18)=   "Label53"
      Tab(1).Control(19)=   "Label54"
      Tab(1).Control(20)=   "Label55"
      Tab(1).Control(21)=   "Label56"
      Tab(1).Control(22)=   "Label57"
      Tab(1).Control(23)=   "meNstp(4)"
      Tab(1).Control(24)=   "meNstp(3)"
      Tab(1).Control(25)=   "meNstp(2)"
      Tab(1).Control(26)=   "meNstp(1)"
      Tab(1).Control(27)=   "meWaiver(4)"
      Tab(1).Control(28)=   "meWaiver(3)"
      Tab(1).Control(29)=   "meWaiver(2)"
      Tab(1).Control(30)=   "meWaiver(1)"
      Tab(1).Control(31)=   "meInternship(4)"
      Tab(1).Control(32)=   "meInternship(3)"
      Tab(1).Control(33)=   "meInternship(2)"
      Tab(1).Control(34)=   "meInternship(1)"
      Tab(1).Control(35)=   "meInternet(4)"
      Tab(1).Control(36)=   "meInternet(3)"
      Tab(1).Control(37)=   "meInternet(2)"
      Tab(1).Control(38)=   "meInternet(1)"
      Tab(1).Control(39)=   "mePowerFee(4)"
      Tab(1).Control(40)=   "mePowerFee(3)"
      Tab(1).Control(41)=   "mePowerFee(2)"
      Tab(1).Control(42)=   "mePowerFee(1)"
      Tab(1).Control(43)=   "meSdf(4)"
      Tab(1).Control(44)=   "meSdf(3)"
      Tab(1).Control(45)=   "meSdf(2)"
      Tab(1).Control(46)=   "meSdf(1)"
      Tab(1).Control(47)=   "meIdNamePlate(4)"
      Tab(1).Control(48)=   "meIdNamePlate(3)"
      Tab(1).Control(49)=   "meIdNamePlate(2)"
      Tab(1).Control(50)=   "meIdNamePlate(1)"
      Tab(1).Control(51)=   "meMta(4)"
      Tab(1).Control(52)=   "meMta(3)"
      Tab(1).Control(53)=   "meMta(2)"
      Tab(1).Control(54)=   "meMta(1)"
      Tab(1).Control(55)=   "meHon(4)"
      Tab(1).Control(56)=   "meHon(3)"
      Tab(1).Control(57)=   "meHon(2)"
      Tab(1).Control(58)=   "meHon(1)"
      Tab(1).Control(59)=   "meRta(4)"
      Tab(1).Control(60)=   "meRta(3)"
      Tab(1).Control(61)=   "meRta(2)"
      Tab(1).Control(62)=   "meRta(1)"
      Tab(1).Control(63)=   "meOjt(4)"
      Tab(1).Control(64)=   "meOjt(3)"
      Tab(1).Control(65)=   "meOjt(2)"
      Tab(1).Control(66)=   "meOjt(1)"
      Tab(1).Control(67)=   "meHrmLab(4)"
      Tab(1).Control(68)=   "meHrmLab(3)"
      Tab(1).Control(69)=   "meHrmLab(2)"
      Tab(1).Control(70)=   "meHrmLab(1)"
      Tab(1).Control(71)=   "meSpeechLab(4)"
      Tab(1).Control(72)=   "meSpeechLab(3)"
      Tab(1).Control(73)=   "meSpeechLab(2)"
      Tab(1).Control(74)=   "meSpeechLab(1)"
      Tab(1).Control(75)=   "meMarineLaboratory(4)"
      Tab(1).Control(76)=   "meMarineLaboratory(3)"
      Tab(1).Control(77)=   "meMarineLaboratory(2)"
      Tab(1).Control(78)=   "meMarineLaboratory(1)"
      Tab(1).Control(79)=   "meNursingAudit(4)"
      Tab(1).Control(80)=   "meNursingAudit(3)"
      Tab(1).Control(81)=   "meNursingAudit(2)"
      Tab(1).Control(82)=   "meNursingAudit(1)"
      Tab(1).Control(83)=   "meAffiliation(4)"
      Tab(1).Control(84)=   "meAffiliation(3)"
      Tab(1).Control(85)=   "meAffiliation(2)"
      Tab(1).Control(86)=   "meAffiliation(1)"
      Tab(1).Control(87)=   "meRle(4)"
      Tab(1).Control(88)=   "meRle(3)"
      Tab(1).Control(89)=   "meRle(2)"
      Tab(1).Control(90)=   "meRle(1)"
      Tab(1).Control(91)=   "meNstp(0)"
      Tab(1).Control(92)=   "meWaiver(0)"
      Tab(1).Control(93)=   "meInternship(0)"
      Tab(1).Control(94)=   "meInternet(0)"
      Tab(1).Control(95)=   "mePowerFee(0)"
      Tab(1).Control(96)=   "meSdf(0)"
      Tab(1).Control(97)=   "meIdNamePlate(0)"
      Tab(1).Control(98)=   "meMta(0)"
      Tab(1).Control(99)=   "meHon(0)"
      Tab(1).Control(100)=   "meRta(0)"
      Tab(1).Control(101)=   "meOjt(0)"
      Tab(1).Control(102)=   "meHrmLab(0)"
      Tab(1).Control(103)=   "meSpeechLab(0)"
      Tab(1).Control(104)=   "meMarineLaboratory(0)"
      Tab(1).Control(105)=   "meNursingAudit(0)"
      Tab(1).Control(106)=   "meAffiliation(0)"
      Tab(1).Control(107)=   "meRle(0)"
      Tab(1).Control(108)=   "txtMiscTotal(0)"
      Tab(1).Control(109)=   "txtMiscTotal(1)"
      Tab(1).Control(110)=   "txtMiscTotal(2)"
      Tab(1).Control(111)=   "txtMiscTotal(3)"
      Tab(1).Control(112)=   "txtMiscTotal(4)"
      Tab(1).ControlCount=   113
      TabCaption(2)   =   "Payment Options"
      TabPicture(2)   =   "frmFees.frx":103B8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(1)=   "Label12"
      Tab(2).Control(2)=   "Label10"
      Tab(2).Control(3)=   "Label4"
      Tab(2).Control(4)=   "Label6"
      Tab(2).Control(5)=   "Label2"
      Tab(2).Control(6)=   "Label5"
      Tab(2).Control(7)=   "Label7"
      Tab(2).Control(8)=   "Label8"
      Tab(2).Control(9)=   "Label9"
      Tab(2).Control(10)=   "txtTotalCashBasis(0)"
      Tab(2).Control(11)=   "txtTotalCashBasis(1)"
      Tab(2).Control(12)=   "txtTotalCashBasis(2)"
      Tab(2).Control(13)=   "txtTotalCashBasis(3)"
      Tab(2).Control(14)=   "txtTotalCashBasis(4)"
      Tab(2).Control(15)=   "txtTotalInstallmentBasis(0)"
      Tab(2).Control(16)=   "txtTotalInstallmentBasis(1)"
      Tab(2).Control(17)=   "txtTotalInstallmentBasis(2)"
      Tab(2).Control(18)=   "txtTotalInstallmentBasis(3)"
      Tab(2).Control(19)=   "txtTotalInstallmentBasis(4)"
      Tab(2).Control(20)=   "txtDownpayment(0)"
      Tab(2).Control(21)=   "txtDownpayment(1)"
      Tab(2).Control(22)=   "txtDownpayment(2)"
      Tab(2).Control(23)=   "txtDownpayment(3)"
      Tab(2).Control(24)=   "txtDownpayment(4)"
      Tab(2).ControlCount=   25
      Begin VB.TextBox txtMiscTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
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
         Height          =   360
         Index           =   4
         Left            =   -66090
         TabIndex        =   199
         Top             =   6030
         Width           =   1455
      End
      Begin VB.TextBox txtMiscTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
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
         Height          =   360
         Index           =   3
         Left            =   -67620
         TabIndex        =   198
         Top             =   6030
         Width           =   1455
      End
      Begin VB.TextBox txtMiscTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
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
         Height          =   360
         Index           =   2
         Left            =   -69150
         TabIndex        =   197
         Top             =   6030
         Width           =   1455
      End
      Begin VB.TextBox txtMiscTotal 
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   -70680
         TabIndex        =   196
         Top             =   6030
         Width           =   1455
      End
      Begin VB.TextBox txtMiscTotal 
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   360
         Index           =   0
         Left            =   -72210
         TabIndex        =   195
         Top             =   6030
         Width           =   1455
      End
      Begin VB.TextBox txtDownpayment 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
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
         Height          =   360
         Index           =   4
         Left            =   -65730
         TabIndex        =   194
         Top             =   2610
         Width           =   1320
      End
      Begin VB.TextBox txtDownpayment 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
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
         Height          =   360
         Index           =   3
         Left            =   -67215
         TabIndex        =   193
         Top             =   2610
         Width           =   1320
      End
      Begin VB.TextBox txtDownpayment 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
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
         Height          =   360
         Index           =   2
         Left            =   -68700
         TabIndex        =   192
         Top             =   2610
         Width           =   1320
      End
      Begin VB.TextBox txtDownpayment 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
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
         Height          =   360
         Index           =   1
         Left            =   -70230
         TabIndex        =   191
         Top             =   2610
         Width           =   1320
      End
      Begin VB.TextBox txtDownpayment 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
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
         Height          =   360
         Index           =   0
         Left            =   -71760
         TabIndex        =   190
         Top             =   2610
         Width           =   1320
      End
      Begin VB.TextBox txtTotalInstallmentBasis 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
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
         Height          =   360
         Index           =   4
         Left            =   -65730
         TabIndex        =   189
         Top             =   1935
         Width           =   1320
      End
      Begin VB.TextBox txtTotalInstallmentBasis 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
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
         Height          =   360
         Index           =   3
         Left            =   -67215
         TabIndex        =   188
         Top             =   1935
         Width           =   1320
      End
      Begin VB.TextBox txtTotalInstallmentBasis 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
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
         Height          =   360
         Index           =   2
         Left            =   -68700
         TabIndex        =   187
         Top             =   1935
         Width           =   1320
      End
      Begin VB.TextBox txtTotalInstallmentBasis 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
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
         Height          =   360
         Index           =   1
         Left            =   -70230
         TabIndex        =   186
         Top             =   1935
         Width           =   1320
      End
      Begin VB.TextBox txtTotalInstallmentBasis 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
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
         Height          =   360
         Index           =   0
         Left            =   -71760
         TabIndex        =   185
         Top             =   1935
         Width           =   1320
      End
      Begin VB.TextBox txtTotalCashBasis 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
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
         Height          =   360
         Index           =   4
         Left            =   -65730
         TabIndex        =   184
         Top             =   1305
         Width           =   1320
      End
      Begin VB.TextBox txtTotalCashBasis 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
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
         Height          =   360
         Index           =   3
         Left            =   -67215
         TabIndex        =   183
         Top             =   1305
         Width           =   1320
      End
      Begin VB.TextBox txtTotalCashBasis 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
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
         Height          =   360
         Index           =   2
         Left            =   -68700
         TabIndex        =   182
         Top             =   1305
         Width           =   1320
      End
      Begin VB.TextBox txtTotalCashBasis 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
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
         Height          =   360
         Index           =   1
         Left            =   -70230
         TabIndex        =   181
         Top             =   1305
         Width           =   1320
      End
      Begin VB.TextBox txtTotalCashBasis 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
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
         Height          =   360
         Index           =   0
         Left            =   -71760
         TabIndex        =   180
         Top             =   1305
         Width           =   1320
      End
      Begin VB.TextBox txtOtherFeesTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00;(#,##0.00)"
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
         ForeColor       =   &H000080FF&
         Height          =   360
         Index           =   4
         Left            =   8370
         TabIndex        =   179
         Top             =   4455
         Width           =   1230
      End
      Begin VB.TextBox txtOtherFeesTotal 
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   360
         Index           =   3
         Left            =   7065
         TabIndex        =   178
         Top             =   4455
         Width           =   1230
      End
      Begin VB.TextBox txtOtherFeesTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
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
         ForeColor       =   &H000080FF&
         Height          =   360
         Index           =   2
         Left            =   5715
         TabIndex        =   177
         Top             =   4455
         Width           =   1230
      End
      Begin VB.TextBox txtOtherFeesTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
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
         ForeColor       =   &H000040C0&
         Height          =   360
         Index           =   1
         Left            =   4365
         TabIndex        =   176
         Top             =   4455
         Width           =   1230
      End
      Begin VB.TextBox txtOtherFeesTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00;(0.00)"
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
         ForeColor       =   &H000080FF&
         Height          =   360
         Index           =   0
         Left            =   3015
         TabIndex        =   175
         Top             =   4455
         Width           =   1230
      End
      Begin MSMask.MaskEdBox meRle 
         Height          =   285
         Index           =   0
         Left            =   -72210
         TabIndex        =   28
         Top             =   675
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meAffiliation 
         Height          =   285
         Index           =   0
         Left            =   -72210
         TabIndex        =   33
         Top             =   990
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meNursingAudit 
         Height          =   285
         Index           =   0
         Left            =   -72210
         TabIndex        =   38
         Top             =   1305
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meMarineLaboratory 
         Height          =   285
         Index           =   0
         Left            =   -72210
         TabIndex        =   43
         Top             =   1620
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meSpeechLab 
         Height          =   285
         Index           =   0
         Left            =   -72210
         TabIndex        =   48
         Top             =   1935
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meHrmLab 
         Height          =   285
         Index           =   0
         Left            =   -72210
         TabIndex        =   53
         Top             =   2250
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meOjt 
         Height          =   285
         Index           =   0
         Left            =   -72210
         TabIndex        =   58
         Top             =   2565
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meRta 
         Height          =   285
         Index           =   0
         Left            =   -72210
         TabIndex        =   63
         Top             =   2880
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meHon 
         Height          =   285
         Index           =   0
         Left            =   -72210
         TabIndex        =   68
         Top             =   3195
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meMta 
         Height          =   285
         Index           =   0
         Left            =   -72210
         TabIndex        =   73
         Top             =   3510
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meIdNamePlate 
         Height          =   285
         Index           =   0
         Left            =   -72210
         TabIndex        =   78
         Top             =   3825
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meSdf 
         Height          =   285
         Index           =   0
         Left            =   -72210
         TabIndex        =   83
         Top             =   4140
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mePowerFee 
         Height          =   285
         Index           =   0
         Left            =   -72210
         TabIndex        =   88
         Top             =   4455
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meInternet 
         Height          =   285
         Index           =   0
         Left            =   -72210
         TabIndex        =   93
         Top             =   4770
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meInternship 
         Height          =   285
         Index           =   0
         Left            =   -72210
         TabIndex        =   98
         Top             =   5085
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meWaiver 
         Height          =   285
         Index           =   0
         Left            =   -72210
         TabIndex        =   103
         Top             =   5400
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meNstp 
         Height          =   285
         Index           =   0
         Left            =   -72210
         TabIndex        =   108
         Top             =   5715
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meRle 
         Height          =   285
         Index           =   1
         Left            =   -70680
         TabIndex        =   29
         Top             =   675
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meRle 
         Height          =   285
         Index           =   2
         Left            =   -69150
         TabIndex        =   30
         Top             =   675
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meRle 
         Height          =   285
         Index           =   3
         Left            =   -67620
         TabIndex        =   31
         Top             =   675
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meRle 
         Height          =   285
         Index           =   4
         Left            =   -66090
         TabIndex        =   32
         Top             =   675
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meAffiliation 
         Height          =   285
         Index           =   1
         Left            =   -70680
         TabIndex        =   34
         Top             =   990
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meAffiliation 
         Height          =   285
         Index           =   2
         Left            =   -69150
         TabIndex        =   35
         Top             =   990
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meAffiliation 
         Height          =   285
         Index           =   3
         Left            =   -67620
         TabIndex        =   36
         Top             =   990
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meAffiliation 
         Height          =   285
         Index           =   4
         Left            =   -66090
         TabIndex        =   37
         Top             =   990
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meNursingAudit 
         Height          =   285
         Index           =   1
         Left            =   -70680
         TabIndex        =   39
         Top             =   1305
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meNursingAudit 
         Height          =   285
         Index           =   2
         Left            =   -69150
         TabIndex        =   40
         Top             =   1305
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meNursingAudit 
         Height          =   285
         Index           =   3
         Left            =   -67620
         TabIndex        =   41
         Top             =   1305
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meNursingAudit 
         Height          =   285
         Index           =   4
         Left            =   -66090
         TabIndex        =   42
         Top             =   1305
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meMarineLaboratory 
         Height          =   285
         Index           =   1
         Left            =   -70680
         TabIndex        =   44
         Top             =   1620
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meMarineLaboratory 
         Height          =   285
         Index           =   2
         Left            =   -69150
         TabIndex        =   45
         Top             =   1620
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meMarineLaboratory 
         Height          =   285
         Index           =   3
         Left            =   -67620
         TabIndex        =   46
         Top             =   1620
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meMarineLaboratory 
         Height          =   285
         Index           =   4
         Left            =   -66090
         TabIndex        =   47
         Top             =   1620
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meSpeechLab 
         Height          =   285
         Index           =   1
         Left            =   -70680
         TabIndex        =   49
         Top             =   1935
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meSpeechLab 
         Height          =   285
         Index           =   2
         Left            =   -69150
         TabIndex        =   50
         Top             =   1935
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meSpeechLab 
         Height          =   285
         Index           =   3
         Left            =   -67620
         TabIndex        =   51
         Top             =   1935
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meSpeechLab 
         Height          =   285
         Index           =   4
         Left            =   -66090
         TabIndex        =   52
         Top             =   1935
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meHrmLab 
         Height          =   285
         Index           =   1
         Left            =   -70680
         TabIndex        =   54
         Top             =   2250
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meHrmLab 
         Height          =   285
         Index           =   2
         Left            =   -69150
         TabIndex        =   55
         Top             =   2250
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meHrmLab 
         Height          =   285
         Index           =   3
         Left            =   -67620
         TabIndex        =   56
         Top             =   2250
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meHrmLab 
         Height          =   285
         Index           =   4
         Left            =   -66090
         TabIndex        =   57
         Top             =   2250
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meOjt 
         Height          =   285
         Index           =   1
         Left            =   -70680
         TabIndex        =   59
         Top             =   2565
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meOjt 
         Height          =   285
         Index           =   2
         Left            =   -69150
         TabIndex        =   60
         Top             =   2565
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meOjt 
         Height          =   285
         Index           =   3
         Left            =   -67620
         TabIndex        =   61
         Top             =   2565
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meOjt 
         Height          =   285
         Index           =   4
         Left            =   -66090
         TabIndex        =   62
         Top             =   2565
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meRta 
         Height          =   285
         Index           =   1
         Left            =   -70680
         TabIndex        =   64
         Top             =   2880
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meRta 
         Height          =   285
         Index           =   2
         Left            =   -69150
         TabIndex        =   65
         Top             =   2880
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meRta 
         Height          =   285
         Index           =   3
         Left            =   -67620
         TabIndex        =   66
         Top             =   2880
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meRta 
         Height          =   285
         Index           =   4
         Left            =   -66090
         TabIndex        =   67
         Top             =   2880
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meHon 
         Height          =   285
         Index           =   1
         Left            =   -70680
         TabIndex        =   69
         Top             =   3195
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meHon 
         Height          =   285
         Index           =   2
         Left            =   -69150
         TabIndex        =   70
         Top             =   3195
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meHon 
         Height          =   285
         Index           =   3
         Left            =   -67620
         TabIndex        =   71
         Top             =   3195
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meHon 
         Height          =   285
         Index           =   4
         Left            =   -66090
         TabIndex        =   72
         Top             =   3195
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meMta 
         Height          =   285
         Index           =   1
         Left            =   -70680
         TabIndex        =   74
         Top             =   3510
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meMta 
         Height          =   285
         Index           =   2
         Left            =   -69150
         TabIndex        =   75
         Top             =   3510
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meMta 
         Height          =   285
         Index           =   3
         Left            =   -67620
         TabIndex        =   76
         Top             =   3510
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meMta 
         Height          =   285
         Index           =   4
         Left            =   -66090
         TabIndex        =   77
         Top             =   3510
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meIdNamePlate 
         Height          =   285
         Index           =   1
         Left            =   -70680
         TabIndex        =   79
         Top             =   3825
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meIdNamePlate 
         Height          =   285
         Index           =   2
         Left            =   -69150
         TabIndex        =   80
         Top             =   3825
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meIdNamePlate 
         Height          =   285
         Index           =   3
         Left            =   -67620
         TabIndex        =   81
         Top             =   3825
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meIdNamePlate 
         Height          =   285
         Index           =   4
         Left            =   -66090
         TabIndex        =   82
         Top             =   3825
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meSdf 
         Height          =   285
         Index           =   1
         Left            =   -70680
         TabIndex        =   84
         Top             =   4140
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meSdf 
         Height          =   285
         Index           =   2
         Left            =   -69150
         TabIndex        =   85
         Top             =   4140
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meSdf 
         Height          =   285
         Index           =   3
         Left            =   -67620
         TabIndex        =   86
         Top             =   4140
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meSdf 
         Height          =   285
         Index           =   4
         Left            =   -66090
         TabIndex        =   87
         Top             =   4140
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mePowerFee 
         Height          =   285
         Index           =   1
         Left            =   -70680
         TabIndex        =   89
         Top             =   4455
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mePowerFee 
         Height          =   285
         Index           =   2
         Left            =   -69150
         TabIndex        =   90
         Top             =   4455
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mePowerFee 
         Height          =   285
         Index           =   3
         Left            =   -67620
         TabIndex        =   91
         Top             =   4455
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mePowerFee 
         Height          =   285
         Index           =   4
         Left            =   -66090
         TabIndex        =   92
         Top             =   4455
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meInternet 
         Height          =   285
         Index           =   1
         Left            =   -70680
         TabIndex        =   94
         Top             =   4770
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meInternet 
         Height          =   285
         Index           =   2
         Left            =   -69150
         TabIndex        =   95
         Top             =   4770
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meInternet 
         Height          =   285
         Index           =   3
         Left            =   -67620
         TabIndex        =   96
         Top             =   4770
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meInternet 
         Height          =   285
         Index           =   4
         Left            =   -66090
         TabIndex        =   97
         Top             =   4770
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meInternship 
         Height          =   285
         Index           =   1
         Left            =   -70680
         TabIndex        =   99
         Top             =   5085
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meInternship 
         Height          =   285
         Index           =   2
         Left            =   -69150
         TabIndex        =   100
         Top             =   5085
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meInternship 
         Height          =   285
         Index           =   3
         Left            =   -67620
         TabIndex        =   101
         Top             =   5085
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meInternship 
         Height          =   285
         Index           =   4
         Left            =   -66090
         TabIndex        =   102
         Top             =   5085
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meWaiver 
         Height          =   285
         Index           =   1
         Left            =   -70680
         TabIndex        =   104
         Top             =   5400
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meWaiver 
         Height          =   285
         Index           =   2
         Left            =   -69150
         TabIndex        =   105
         Top             =   5400
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meWaiver 
         Height          =   285
         Index           =   3
         Left            =   -67620
         TabIndex        =   106
         Top             =   5400
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meWaiver 
         Height          =   285
         Index           =   4
         Left            =   -66090
         TabIndex        =   107
         Top             =   5400
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meNstp 
         Height          =   285
         Index           =   1
         Left            =   -70680
         TabIndex        =   109
         Top             =   5715
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meNstp 
         Height          =   285
         Index           =   2
         Left            =   -69150
         TabIndex        =   110
         Top             =   5715
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meNstp 
         Height          =   285
         Index           =   3
         Left            =   -67620
         TabIndex        =   111
         Top             =   5715
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meNstp 
         Height          =   285
         Index           =   4
         Left            =   -66090
         TabIndex        =   112
         Top             =   5715
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meEntrance 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3015
         TabIndex        =   1
         Top             =   1125
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##,##0.00;(##,##0.00)"
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meTuitionFee 
         Height          =   285
         Index           =   0
         Left            =   3015
         TabIndex        =   2
         Top             =   1575
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meTuitionFee 
         Height          =   285
         Index           =   1
         Left            =   4365
         TabIndex        =   3
         Top             =   1575
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meTuitionFee 
         Height          =   285
         Index           =   2
         Left            =   5715
         TabIndex        =   4
         Top             =   1575
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meTuitionFee 
         Height          =   285
         Index           =   3
         Left            =   7065
         TabIndex        =   5
         Top             =   1575
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meTuitionFee 
         Height          =   285
         Index           =   4
         Left            =   8370
         TabIndex        =   6
         Top             =   1575
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meRegistration 
         Height          =   285
         Index           =   0
         Left            =   3015
         TabIndex        =   7
         Top             =   2385
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meLibrary 
         Height          =   285
         Index           =   0
         Left            =   3015
         TabIndex        =   8
         Top             =   2790
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meLibrary 
         Height          =   285
         Index           =   1
         Left            =   4365
         TabIndex        =   9
         Top             =   2790
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meLibrary 
         Height          =   285
         Index           =   2
         Left            =   5715
         TabIndex        =   10
         Top             =   2790
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meLibrary 
         Height          =   285
         Index           =   3
         Left            =   7065
         TabIndex        =   11
         Top             =   2790
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meLibrary 
         Height          =   285
         Index           =   4
         Left            =   8370
         TabIndex        =   12
         Top             =   2790
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meLaboratory 
         Height          =   285
         Index           =   0
         Left            =   3015
         TabIndex        =   13
         Top             =   3195
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meAthleticFee 
         Height          =   285
         Index           =   0
         Left            =   3015
         TabIndex        =   18
         Top             =   3600
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meGuidanceAndCounselor 
         Height          =   285
         Index           =   0
         Left            =   3015
         TabIndex        =   23
         Top             =   4005
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meLaboratory 
         Height          =   285
         Index           =   1
         Left            =   4365
         TabIndex        =   14
         Top             =   3195
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meLaboratory 
         Height          =   285
         Index           =   2
         Left            =   5715
         TabIndex        =   15
         Top             =   3195
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meLaboratory 
         Height          =   285
         Index           =   3
         Left            =   7065
         TabIndex        =   16
         Top             =   3195
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meLaboratory 
         Height          =   285
         Index           =   4
         Left            =   8370
         TabIndex        =   17
         Top             =   3195
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meAthleticFee 
         Height          =   285
         Index           =   1
         Left            =   4365
         TabIndex        =   19
         Top             =   3600
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meAthleticFee 
         Height          =   285
         Index           =   2
         Left            =   5715
         TabIndex        =   20
         Top             =   3600
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meAthleticFee 
         Height          =   285
         Index           =   3
         Left            =   7065
         TabIndex        =   21
         Top             =   3600
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meAthleticFee 
         Height          =   285
         Index           =   4
         Left            =   8370
         TabIndex        =   22
         Top             =   3600
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meGuidanceAndCounselor 
         Height          =   285
         Index           =   1
         Left            =   4365
         TabIndex        =   24
         Top             =   4005
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meGuidanceAndCounselor 
         Height          =   285
         Index           =   2
         Left            =   5715
         TabIndex        =   25
         Top             =   4005
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meGuidanceAndCounselor 
         Height          =   285
         Index           =   3
         Left            =   7065
         TabIndex        =   26
         Top             =   4005
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meGuidanceAndCounselor 
         Height          =   285
         Index           =   4
         Left            =   8370
         TabIndex        =   27
         Top             =   4005
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meEntrance 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   4365
         TabIndex        =   200
         Top             =   1125
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##,##0.00;(##,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meEntrance 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   5715
         TabIndex        =   201
         Top             =   1125
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##,##0.00;(##,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meEntrance 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   7065
         TabIndex        =   202
         Top             =   1125
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##,##0.00;(##,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meEntrance 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   8370
         TabIndex        =   203
         Top             =   1125
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##,##0.00;(##,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meRegistration 
         Height          =   285
         Index           =   1
         Left            =   4365
         TabIndex        =   204
         Top             =   2385
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meRegistration 
         Height          =   285
         Index           =   2
         Left            =   5715
         TabIndex        =   205
         Top             =   2385
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meRegistration 
         Height          =   285
         Index           =   3
         Left            =   7065
         TabIndex        =   206
         Top             =   2385
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meRegistration 
         Height          =   285
         Index           =   4
         Left            =   8370
         TabIndex        =   207
         Top             =   2385
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###,###.##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "1st Year"
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
         Left            =   -71535
         TabIndex        =   174
         Top             =   945
         Width           =   960
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "2nd Year"
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
         Left            =   -69915
         TabIndex        =   173
         Top             =   945
         Width           =   960
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "3rd Year"
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
         Left            =   -68430
         TabIndex        =   172
         Top             =   945
         Width           =   960
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "4th Year"
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
         Left            =   -66900
         TabIndex        =   171
         Top             =   945
         Width           =   960
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "5th Year"
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
         Left            =   -65415
         TabIndex        =   170
         Top             =   945
         Width           =   960
      End
      Begin VB.Label Label63 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   495
         TabIndex        =   169
         Top             =   4410
         Width           =   1230
      End
      Begin VB.Label Label62 
         BackStyle       =   0  'Transparent
         Caption         =   "Guidance and Counselor"
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
         Left            =   495
         TabIndex        =   168
         Top             =   4005
         Width           =   2175
      End
      Begin VB.Label Label61 
         BackStyle       =   0  'Transparent
         Caption         =   "Athletic Fee"
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
         Left            =   495
         TabIndex        =   167
         Top             =   3600
         Width           =   2085
      End
      Begin VB.Label Label60 
         BackStyle       =   0  'Transparent
         Caption         =   "Laboratory *"
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
         Left            =   495
         TabIndex        =   166
         Top             =   3195
         Width           =   1230
      End
      Begin VB.Label Label59 
         BackStyle       =   0  'Transparent
         Caption         =   "Library *"
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
         Left            =   495
         TabIndex        =   165
         Top             =   2790
         Width           =   1230
      End
      Begin VB.Label Label58 
         BackStyle       =   0  'Transparent
         Caption         =   "Registration"
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
         Left            =   495
         TabIndex        =   164
         Top             =   2430
         Width           =   1230
      End
      Begin VB.Label Label57 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74370
         TabIndex        =   163
         Top             =   6075
         Width           =   1950
      End
      Begin VB.Label Label56 
         BackStyle       =   0  'Transparent
         Caption         =   "Nstp"
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
         Left            =   -74325
         TabIndex        =   162
         Top             =   5760
         Width           =   2130
      End
      Begin VB.Label Label55 
         BackStyle       =   0  'Transparent
         Caption         =   "Waiver"
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
         Left            =   -74325
         TabIndex        =   161
         Top             =   5445
         Width           =   2130
      End
      Begin VB.Label Label54 
         BackStyle       =   0  'Transparent
         Caption         =   "Internship"
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
         Left            =   -74325
         TabIndex        =   160
         Top             =   5130
         Width           =   2130
      End
      Begin VB.Label Label53 
         BackStyle       =   0  'Transparent
         Caption         =   "Internet"
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
         Left            =   -74325
         TabIndex        =   159
         Top             =   4815
         Width           =   2130
      End
      Begin VB.Label Label52 
         BackStyle       =   0  'Transparent
         Caption         =   "Power Fee"
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
         Left            =   -74325
         TabIndex        =   158
         Top             =   4500
         Width           =   2130
      End
      Begin VB.Label Label51 
         BackStyle       =   0  'Transparent
         Caption         =   "Sdf"
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
         Left            =   -74325
         TabIndex        =   157
         Top             =   4185
         Width           =   2130
      End
      Begin VB.Label Label50 
         BackStyle       =   0  'Transparent
         Caption         =   "Id/Name Plate"
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
         Left            =   -74325
         TabIndex        =   156
         Top             =   3870
         Width           =   2130
      End
      Begin VB.Label Label49 
         BackStyle       =   0  'Transparent
         Caption         =   "Mta"
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
         Left            =   -74325
         TabIndex        =   155
         Top             =   3555
         Width           =   2130
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         Caption         =   "H-On"
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
         Left            =   -74325
         TabIndex        =   154
         Top             =   3240
         Width           =   2130
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "Rta"
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
         Left            =   -74325
         TabIndex        =   153
         Top             =   2925
         Width           =   2130
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "Ojt"
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
         Left            =   -74325
         TabIndex        =   152
         Top             =   2610
         Width           =   2130
      End
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
         Caption         =   "Hrm Lab"
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
         Left            =   -74325
         TabIndex        =   151
         Top             =   2295
         Width           =   2130
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Speech Lab"
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
         Left            =   -74325
         TabIndex        =   150
         Top             =   1980
         Width           =   2130
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "Marine Laboratory *"
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
         Left            =   -74325
         TabIndex        =   149
         Top             =   1665
         Width           =   2130
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Nursing Audit"
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
         Left            =   -74325
         TabIndex        =   148
         Top             =   1350
         Width           =   2130
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Affiliation"
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
         Left            =   -74325
         TabIndex        =   147
         Top             =   1035
         Width           =   2130
      End
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   "5th Year"
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
         Left            =   -65775
         TabIndex        =   146
         Top             =   405
         Width           =   960
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "4th Year"
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
         Left            =   -67260
         TabIndex        =   145
         Top             =   405
         Width           =   960
      End
      Begin VB.Label Label40 
         BackStyle       =   0  'Transparent
         Caption         =   "3rd Year"
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
         Left            =   -68790
         TabIndex        =   144
         Top             =   405
         Width           =   960
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "2nd Year"
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
         Left            =   -70275
         TabIndex        =   143
         Top             =   405
         Width           =   960
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "1st Year"
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
         Left            =   -71895
         TabIndex        =   142
         Top             =   405
         Width           =   960
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "RLE"
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
         Left            =   -74325
         TabIndex        =   141
         Top             =   720
         Width           =   2130
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "5th Year"
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
         Left            =   8505
         TabIndex        =   140
         Top             =   630
         Width           =   960
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "4th Year"
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
         Left            =   7155
         TabIndex        =   139
         Top             =   630
         Width           =   960
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "3rd Year"
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
         Left            =   5850
         TabIndex        =   138
         Top             =   630
         Width           =   960
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "2nd Year"
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
         Left            =   4455
         TabIndex        =   137
         Top             =   630
         Width           =   960
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "1st Year"
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
         Left            =   3285
         TabIndex        =   136
         Top             =   630
         Width           =   960
      End
      Begin VB.Label Label29 
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
         TabIndex        =   135
         Top             =   585
         Width           =   1275
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Breakdown"
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
         Left            =   225
         TabIndex        =   134
         Top             =   540
         Width           =   1635
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "OTHER FEES:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   128
         Top             =   2070
         Width           =   1230
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Tuition Fee *"
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
         Left            =   450
         TabIndex        =   127
         Top             =   1620
         Width           =   1230
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Entrance"
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
         Left            =   450
         TabIndex        =   126
         Top             =   1170
         Width           =   1230
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Down Payment"
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
         Left            =   -74550
         TabIndex        =   125
         Top             =   2655
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Installment Basis"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74595
         TabIndex        =   124
         Top             =   1980
         Width           =   2580
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment"
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
         Left            =   -74760
         TabIndex        =   123
         Top             =   390
         Width           =   1275
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "options"
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
         Left            =   -73410
         TabIndex        =   122
         Top             =   435
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Cash Basis"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -74595
         TabIndex        =   121
         Top             =   1350
         Width           =   2625
      End
   End
   Begin VB.PictureBox picFooter 
      Align           =   2  'Align Bottom
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   11160
      TabIndex        =   118
      Top             =   7800
      Width           =   11220
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "Close"
         Height          =   315
         Left            =   9915
         TabIndex        =   115
         Top             =   60
         Width           =   1035
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H000080FF&
         Caption         =   "Delete"
         Height          =   315
         Left            =   8835
         MaskColor       =   &H000080FF&
         Style           =   1  'Graphical
         TabIndex        =   114
         Top             =   60
         Width           =   1035
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   315
         Left            =   7755
         TabIndex        =   113
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
         TabIndex        =   119
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
      ScaleWidth      =   11220
      TabIndex        =   116
      Top             =   0
      Width           =   11220
      Begin VB.Image Image2 
         Height          =   375
         Left            =   10845
         Picture         =   "frmFees.frx":103D4
         Top             =   0
         Width           =   390
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fees Schedule"
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
         TabIndex        =   117
         Top             =   120
         Width           =   4260
      End
      Begin VB.Image Image1 
         Height          =   780
         Left            =   0
         Picture         =   "frmFees.frx":10884
         Top             =   60
         Width           =   915
      End
   End
   Begin VB.Label Label24 
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
      Height          =   285
      Left            =   225
      TabIndex        =   133
      Top             =   1035
      Width           =   1230
   End
   Begin VB.Label Label23 
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
      Left            =   1350
      TabIndex        =   132
      Top             =   45
      Width           =   1275
   End
   Begin VB.Label Label22 
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
      Left            =   0
      TabIndex        =   131
      Top             =   0
      Width           =   1275
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
      TabIndex        =   130
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
      TabIndex        =   129
      Top             =   0
      Width           =   1230
   End
End
Attribute VB_Name = "frmFees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsData As New ADODB.Recordset

Private Sub dcCourses_Change()
    Dim strMask As String
    ValidateAccessLevel Me, "Insert"
    'Dim temp As MSMask.MaskEdBox
    ClearEntries Me
    
    'retrieve fees
    strSql = "SELECT * FROM dbo.Fees WHERE CourseCode = '" & dcCourses.BoundText & "' AND CurrentSyId = " & SchoolInformation.CurrentSyId & " AND CurrentSemesterId = " & SchoolInformation.CurrentSemesterId & " ORDER BY CurrentYearLevelId;"
    'Debug.Print strSql
    Set rsData = GetRecordset(strSql)
    If rsData.State = adStateOpen Then
        If rsData.RecordCount > 0 Then
            ValidateAccessLevel Me, "Update"
            lblMessage.Caption = "Retrieving records.... pls wait..."
            Dim x As Integer
            x = 0
            Do While Not rsData.EOF And x < 5
                strMask = meEntrance(x).Mask
                meEntrance(x).Mask = ""
                meEntrance(x).Text = ""
                meEntrance(x) = Format$(rsData!Entrance, "###,##0.00")
                'meEntrance(x).Mask = strMask
                
                strMask = meTuitionFee(x).Mask
                meTuitionFee(x).Mask = ""
                meTuitionFee(x) = Format$(rsData!TuitionFee, "###,##0.00")
                'meTuitionFee(x).Mask = strMask
                
                strMask = meRegistration(x).Mask
                meRegistration(x).Mask = ""
                meRegistration(x) = Format$(rsData!Registration, "###,##0.00")
                'meRegistration(x).Mask = strMask
                
                strMask = meLibrary(x).Mask
                meLibrary(x).Mask = ""
                meLibrary(x) = Format$(rsData!Library, "###,##0.00")
                'meLibrary(x).Mask = strMask
                
                strMask = meLaboratory(x).Mask
                meLaboratory(x).Mask = ""
                meLaboratory(x) = Format$(rsData!Laboratory, "###,##0.00")
                'meLaboratory(x).Mask = strMask
                
                strMask = meAthleticFee(x).Mask
                meAthleticFee(x).Mask = ""
                meAthleticFee(x) = Format$(rsData!AthleticFee, "###,##0.00")
                'meAthleticFee(x).Mask = strMask
                
                strMask = meGuidanceAndCounselor(x).Mask
                meGuidanceAndCounselor(x).Mask = ""
                meGuidanceAndCounselor(x) = Format$(rsData!GuidanceAndCounselor, "###,##0.00")
                'meGuidanceAndCounselor(x).Mask = strMask
                                
                'misc fees
                strMask = meRle(x).Mask
                meRle(x).Mask = ""
                meRle(x) = Format$(rsData!Rle, "###,##0.00")
                'meRle(x).Mask = strMask
                
                strMask = meAffiliation(x).Mask
                meAffiliation(x).Mask = ""
                meAffiliation(x) = Format$(rsData!Affiliation, "###,##0.00")
                'meAffiliation(x).Mask = strMask
                
                strMask = meNursingAudit(x).Mask
                meNursingAudit(x).Mask = ""
                meNursingAudit(x) = Format$(rsData!NursingAudit, "###,##0.00")
                'meNursingAudit(x).Mask = strMask
                
                strMask = meMarineLaboratory(x).Mask
                meMarineLaboratory(x).Mask = ""
                meMarineLaboratory(x) = Format$(rsData!MarineLaboratory, "###,##0.00")
                'meMarineLaboratory(x).Mask = strMask
                
                strMask = meSpeechLab(x).Mask
                meSpeechLab(x).Mask = ""
                meSpeechLab(x) = Format$(rsData!SpeechLab, "###,##0.00")
                'meSpeechLab(x).Mask = strMask
                
                strMask = meHrmLab(x).Mask
                meHrmLab(x).Mask = ""
                meHrmLab(x) = Format$(rsData!HrmLab, "###,##0.00")
                'meHrmLab(x).Mask = strMask
                
                strMask = meOjt(x).Mask
                meOjt(x).Mask = ""
                meOjt(x) = Format$(rsData!Ojt, "###,##0.00")
                'meOjt(x).Mask = strMask
                
                strMask = meRta(x).Mask
                meRta(x).Mask = ""
                meRta(x) = Format$(rsData!Rta, "###,##0.00")
                'meRta(x).Mask = strMask
                
                strMask = meHon(x).Mask
                meHon(x).Mask = ""
                meHon(x) = Format$(rsData!HOn, "###,##0.00")
                'meHon(x).Mask = strMask
                
                strMask = meMta(x).Mask
                meMta(x).Mask = ""
                meMta(x) = Format$(rsData!Mta, "###,##0.00")
                'meMta(x).Mask = strMask
                
                strMask = meIdNamePlate(x).Mask
                meIdNamePlate(x).Mask = ""
                meIdNamePlate(x) = Format$(rsData!IdNamePlate, "###,##0.00")
                'meIdNamePlate(x).Mask = strMask
                
                strMask = meSdf(x).Mask
                meSdf(x).Mask = ""
                meSdf(x) = Format$(rsData!Sdf, "###,##0.00")
                'meSdf(x).Mask = strMask
                
                strMask = mePowerFee(x).Mask
                mePowerFee(x).Mask = ""
                mePowerFee(x) = Format$(rsData!PowerFee, "###,##0.00")
                'mePowerFee(x).Mask = strMask
                
                strMask = meInternet(x).Mask
                meInternet(x).Mask = ""
                meInternet(x) = Format$(rsData!Internet, "###,##0.00")
                'meInternet(x).Mask = strMask
                
                strMask = meInternship(x).Mask
                meInternship(x).Mask = ""
                meInternship(x) = Format$(rsData!Internship, "###,##0.00")
                'meInternship(x).Mask = strMask
                
                strMask = meWaiver(x).Mask
                meWaiver(x).Mask = ""
                meWaiver(x) = Format$(rsData!Waiver, "###,##0.00")
                'meWaiver(x).Mask = strMask
                
                strMask = meNstp(x).Mask
                meNstp(x).Mask = ""
                meNstp(x) = Format$(rsData!Nstp, "###,##0.00")
                'meNstp(x).Mask = strMask
                
                rsData.MoveNext
                
                x = x + 1
            Loop
            lblMessage.Caption = "Calculating.... pls wait..."
            ComputeTotal
        End If
    End If
    lblMessage.Caption = ""
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If cmdSave.Caption = "Update" Then
        If MsgBox("Do you want to Delete the record?", vbYesNo + vbCritical, Me.Caption) = vbYes Then
            strSql = "UPDATE dbo.Fees SET  Deleted = 1 " & _
                        "WHERE CourseCode = '" & dcCourses.BoundText & "'  AND " & _
                                             "CurrentSyId = " & SchoolInformation.CurrentSyId & " AND " & _
                                             "CurrentSemesterId = " & SchoolInformation.CurrentSemesterId & "; "
            'Execute SQL Command
            RunSql (strSql)
            

            SSTab1.Tab = 0
            lblMessage.Caption = ""
            ClearEntries Me
            dcCourses.SetFocus
        End If
    End If
End Sub

Private Sub cmdSave_Click()
    Dim x As Integer
    If EntriesValid Then
        If MsgBox("Do you want to " & cmdSave.Caption & " record?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
            lblMessage.Caption = "Saving information... pls wait..."
            strSql = ""
            
            If cmdSave.Caption <> "Update" Then
                For x = 0 To 4 Step 1
                    strSql = strSql + "INSERT INTO dbo.Fees(CourseCode, CurrentSyId, CurrentSemesterId, CurrentYearLevelId, " & _
                                                      "Entrance, TuitionFee, Registration, Library, Laboratory, AthleticFee, GuidanceAndCounselor, " & _
                                                      "Rle, NursingAudit, MarineLaboratory, SpeechLab, HrmLab, Ojt, Rta, HOn, Mta, IdNamePlate, " & _
                                                      "Sdf, PowerFee, Internet, Internship, Waiver, Nstp, Deleted) " & _
                                            "VALUES('" & dcCourses.BoundText & "', " & _
                                                   "" & SchoolInformation.CurrentSyId & ", " & _
                                                   "" & SchoolInformation.CurrentSemesterId & ", " & _
                                                   "" & Trim(Str(x + 1)) & ", " & _
                                                   Val(Replace(Replace(meEntrance(x), ",", ""), "_", "")) & ", " & _
                                                   Val(Replace(Replace(meTuitionFee(x), ",", ""), "_", "")) & ", " & _
                                                   Val(Replace(Replace(meRegistration(x), ",", ""), "_", "")) & ", " & _
                                                   Val(Replace(Replace(meLibrary(x), ",", ""), "_", "")) & ", " & _
                                                   Val(Replace(Replace(meLaboratory(x), ",", ""), "_", "")) & ", " & _
                                                   Val(Replace(Replace(meAthleticFee(x), ",", ""), "_", "")) & ", " & _
                                                   Val(Replace(Replace(meGuidanceAndCounselor(x), ",", ""), "_", "")) & ", " & _
                                                   Val(Replace(Replace(meRle(x), ",", ""), "_", "")) & ", " & _
                                                   Val(Replace(Replace(meNursingAudit(x), ",", ""), "_", "")) & ", " & _
                                                   Val(Replace(Replace(meMarineLaboratory(x), ",", ""), "_", "")) & ", " & _
                                                   Val(Replace(Replace(meSpeechLab(x), ",", ""), "_", "")) & ", " & _
                                                   Val(Replace(Replace(meHrmLab(x), ",", ""), "_", "")) & ", " & _
                                                   Val(Replace(Replace(meOjt(x), ",", ""), "_", "")) & ", " & _
                                                   Val(Replace(Replace(meRta(x), ",", ""), "_", "")) & ", " & _
                                                   Val(Replace(Replace(meHon(x), ",", ""), "_", "")) & ", " & _
                                                   Val(Replace(Replace(meMta(x), ",", ""), "_", "")) & ", " & _
                                                   Val(Replace(Replace(meIdNamePlate(x), ",", ""), "_", "")) & ","
                    strSql = strSql + " " & Val(Replace(Replace(meSdf(x), ",", ""), "_", "")) & ", " & _
                                      Val(Replace(Replace(mePowerFee(x), ",", ""), "_", "")) & ", " & _
                                      Val(Replace(Replace(meInternet(x), ",", ""), "_", "")) & ", " & _
                                      Val(Replace(Replace(meInternship(x), ",", ""), "_", "")) & ", " & _
                                      Val(Replace(Replace(meWaiver(x), ",", ""), "_", "")) & ", " & _
                                      Val(Replace(Replace(meNstp(x), ",", ""), "_", "")) & ", 0); "
                Next x

            Else
                For x = 0 To 4 Step 1
                    strSql = strSql + "UPDATE dbo.Fees SET Entrance = " & Val(Replace(Replace(meEntrance(x), ",", ""), "_", "")) & ", " & _
                                                          "TuitionFee = " & Val(Replace(Replace(meTuitionFee(x), ",", ""), "_", "")) & ", " & _
                                                          "Registration = " & Val(Replace(Replace(meRegistration(x), ",", ""), "_", "")) & ", " & _
                                                          "Library = " & Val(Replace(Replace(meLibrary(x), ",", ""), "_", "")) & ", " & _
                                                          "Laboratory = " & Val(Replace(Replace(meLaboratory(x), ",", ""), "_", "")) & ", " & _
                                                          "AthleticFee = " & Val(Replace(Replace(meAthleticFee(x), ",", ""), "_", "")) & ", " & _
                                                          "GuidanceAndCounselor = " & Val(Replace(Replace(meGuidanceAndCounselor(x), ",", ""), "_", "")) & ", "
                    strSql = strSql + "Rle = " & Val(Replace(Replace(meRle(x), ",", ""), "_", "")) & ", " & _
                                                          "NursingAudit = " & Val(Replace(Replace(meNursingAudit(x), ",", ""), "_", "")) & ", " & _
                                                          "MarineLaboratory = " & Val(Replace(Replace(meMarineLaboratory(x), ",", ""), "_", "")) & ", " & _
                                                          "SpeechLab = " & Val(Replace(Replace(meSpeechLab(x), ",", ""), "_", "")) & ", " & _
                                                          "HrmLab = " & Val(Replace(Replace(meHrmLab(x), ",", ""), "_", "")) & ", " & _
                                                          "Ojt = " & Val(Replace(Replace(meOjt(x), ",", ""), "_", "")) & ", " & _
                                                          "Rta = " & Val(Replace(Replace(meRta(x), ",", ""), "_", "")) & ", " & _
                                                          "HOn = " & Val(Replace(Replace(meHon(x), ",", ""), "_", "")) & ", " & _
                                                          "Mta = " & Val(Replace(Replace(meMta(x), ",", ""), "_", "")) & ", " & _
                                                          "IdNamePlate = " & Val(Replace(Replace(meIdNamePlate(x), ",", ""), "_", "")) & ", " & _
                                                          "Sdf = " & Val(Replace(Replace(meSdf(x), ",", ""), "_", "")) & ", " & _
                                                          "PowerFee = " & Val(Replace(Replace(mePowerFee(x), ",", ""), "_", "")) & ", " & _
                                                          "Internet = " & Val(Replace(Replace(meInternet(x), ",", ""), "_", "")) & ", " & _
                                                          "Internship = " & Val(Replace(Replace(meInternship(x), ",", ""), "_", "")) & ", " & _
                                                          "Waiver = " & Val(Replace(Replace(meWaiver(x), ",", ""), "_", "")) & ", " & _
                                                          "Nstp = " & Val(Replace(Replace(meNstp(x), ",", ""), "_", "")) & ", " & _
                                                          "Deleted = 0 "
                    strSql = strSql + " WHERE CourseCode = '" & dcCourses.BoundText & "'  AND " & _
                                             "CurrentSyId = " & SchoolInformation.CurrentSyId & " AND " & _
                                             "CurrentSemesterId = " & SchoolInformation.CurrentSemesterId & " AND " & _
                                             "CurrentYearLevelId = " & Trim(Str(x + 1)) & "; "
                Next x
            End If
            RunSql (strSql)

            lblMessage.Caption = ""
            ClearEntries Me
            dcCourses.SetFocus
        End If
    End If
End Sub



Private Sub Form_Load()
    SSTab1.Tab = 0
    BindDataCombo "SELECT * FROM dbo.Courses", "CourseDesc", dcCourses, "CourseCode", True 'Bind Courses
End Sub

Private Sub Form_Activate()
    'MakeTransparent Me.hwnd, 255 'Unfade form
End Sub

Private Sub Form_Deactivate()
    'MakeTransparent Me.hwnd, 200 'Fade Form
End Sub

Private Function EntriesValid() As Boolean
    EntriesValid = False
    If dcCourses.Tag = "" Then
        MsgBox ("Course required!")
        dcCourses.SetFocus
        Exit Function
    End If
    EntriesValid = True
End Function

Private Sub ComputeTotal()
    Dim x As Integer
    
    'entrance
    txtOtherFeesTotal(0).Text = Format(Val(Replace(Replace(meEntrance(0), ",", ""), "_", "")), "#,###,##0.00")
    txtOtherFeesTotal(1).Text = Format(0, "#,###,###.##")
    txtOtherFeesTotal(2).Text = Format(0, "#,###,###.##")
    txtOtherFeesTotal(3).Text = Format(0, "#,###,###.##")
    txtOtherFeesTotal(4).Text = Format(0, "#,###,###.##")
    'tuition fees
    For x = 0 To 4 Step 1
        txtOtherFeesTotal(x).Text = Format(Val(Replace(txtOtherFeesTotal(x).Text, ",", "")) + Val(Replace(Replace(meTuitionFee(x), ",", ""), "_", "")), "#,###,##0.00")
    Next x
    'registration
    For x = 0 To 4 Step 1
        txtOtherFeesTotal(x).Text = Format(Val(Replace(txtOtherFeesTotal(x).Text, ",", "")) + Val(Replace(Replace(meRegistration(x), ",", ""), "_", "")), "#,###,##0.00")
    Next x
    'txtOtherFeesTotal(0).Text = Format(Val(Replace(txtOtherFeesTotal(0).Text, ",", "")) + Val(Replace(Replace(meRegistration(0), ",", ""), "_", "")), "#,###,##0.00")
    'library
    For x = 0 To 4 Step 1
        txtOtherFeesTotal(x).Text = Format(Val(Replace(txtOtherFeesTotal(x).Text, ",", "")) + Val(Replace(Replace(meLibrary(x), ",", ""), "_", "")), "#,###,##0.00")
    Next x
    'laboratory
    For x = 0 To 4 Step 1
        txtOtherFeesTotal(x).Text = Format(Val(Replace(txtOtherFeesTotal(x).Text, ",", "")) + Val(Replace(Replace(meLaboratory(x), ",", ""), "_", "")), "#,###,##0.00")
    Next x
    'Athletic Fee
    For x = 0 To 4 Step 1
        txtOtherFeesTotal(x).Text = Format(Val(Replace(txtOtherFeesTotal(x).Text, ",", "")) + Val(Replace(Replace(meAthleticFee(x), ",", ""), "_", "")), "#,###,##0.00")
    Next x
    'GuidanceAndCounselor
    For x = 0 To 4 Step 1
        txtOtherFeesTotal(x).Text = Format(Val(Replace(txtOtherFeesTotal(x).Text, ",", "")) + Val(Replace(Replace(meGuidanceAndCounselor(x), ",", ""), "_", "")), "#,###,##0.00")
    Next x
    
    'misc fees
    'rle
    For x = 0 To 4 Step 1
        txtMiscTotal(x).Text = Format(Val(Replace(Replace(meRle(x), ",", ""), "_", "")), "#,###,##0.00")
    Next x
    'affiliation
    For x = 0 To 4 Step 1
        txtMiscTotal(x).Text = Format(Val(Replace(txtMiscTotal(x).Text, ",", "")) + Val(Replace(Replace(meAffiliation(x), ",", ""), "_", "")), "#,###,##0.00")
    Next x
    'nursing audit
    For x = 0 To 4 Step 1
        txtMiscTotal(x).Text = Format(Val(Replace(txtMiscTotal(x).Text, ",", "")) + Val(Replace(Replace(meNursingAudit(x), ",", ""), "_", "")), "#,###,##0.00")
    Next x
    'Marine Laboratory
    For x = 0 To 4 Step 1
        txtMiscTotal(x).Text = Format(Val(Replace(txtMiscTotal(x).Text, ",", "")) + Val(Replace(Replace(meMarineLaboratory(x), ",", ""), "_", "")), "#,###,###.##")
    Next x
    'meSpeechLab
    For x = 0 To 4 Step 1
        txtMiscTotal(x).Text = Format(Val(Replace(txtMiscTotal(x).Text, ",", "")) + Val(Replace(Replace(meSpeechLab(x), ",", ""), "_", "")), "#,###,###.##")
    Next x
    'meHrmLab
    For x = 0 To 4 Step 1
        txtMiscTotal(x).Text = Format(Val(Replace(txtMiscTotal(x).Text, ",", "")) + Val(Replace(Replace(meHrmLab(x), ",", ""), "_", "")), "#,###,###.##")
    Next x
    'meOjt
    For x = 0 To 4 Step 1
        txtMiscTotal(x).Text = Format(Val(Replace(txtMiscTotal(x).Text, ",", "")) + Val(Replace(Replace(meOjt(x), ",", ""), "_", "")), "#,###,###.##")
    Next x
    'meRta
    For x = 0 To 4 Step 1
        txtMiscTotal(x).Text = Format(Val(Replace(txtMiscTotal(x).Text, ",", "")) + Val(Replace(Replace(meRta(x), ",", ""), "_", "")), "#,###,###.##")
    Next x
    'meHon
    For x = 0 To 4 Step 1
        txtMiscTotal(x).Text = Format(Val(Replace(txtMiscTotal(x).Text, ",", "")) + Val(Replace(Replace(meHon(x), ",", ""), "_", "")), "#,###,###.##")
    Next x
    'meMta
    For x = 0 To 4 Step 1
        txtMiscTotal(x).Text = Format(Val(Replace(txtMiscTotal(x).Text, ",", "")) + Val(Replace(Replace(meMta(x), ",", ""), "_", "")), "#,###,###.##")
    Next x
    'meIdNamePlate
    For x = 0 To 4 Step 1
        txtMiscTotal(x).Text = Format(Val(Replace(txtMiscTotal(x).Text, ",", "")) + Val(Replace(Replace(meIdNamePlate(x), ",", ""), "_", "")), "#,###,###.##")
    Next x
    'meSdf
    For x = 0 To 4 Step 1
        txtMiscTotal(x).Text = Format(Val(Replace(txtMiscTotal(x).Text, ",", "")) + Val(Replace(Replace(meSdf(x), ",", ""), "_", "")), "#,###,###.##")
    Next x
    'mePowerFee
    For x = 0 To 4 Step 1
        txtMiscTotal(x).Text = Format(Val(Replace(txtMiscTotal(x).Text, ",", "")) + Val(Replace(Replace(mePowerFee(x), ",", ""), "_", "")), "#,###,###.##")
    Next x
    'meInternet
    For x = 0 To 4 Step 1
        txtMiscTotal(x).Text = Format(Val(Replace(txtMiscTotal(x).Text, ",", "")) + Val(Replace(Replace(meInternet(x), ",", ""), "_", "")), "#,###,###.##")
    Next x
    'meInternship
    For x = 0 To 4 Step 1
        txtMiscTotal(x).Text = Format(Val(Replace(txtMiscTotal(x).Text, ",", "")) + Val(Replace(Replace(meInternship(x), ",", ""), "_", "")), "#,###,###.##")
    Next x
    'meWaiver
    For x = 0 To 4 Step 1
        txtMiscTotal(x).Text = Format(Val(Replace(txtMiscTotal(x).Text, ",", "")) + Val(Replace(Replace(meWaiver(x), ",", ""), "_", "")), "#,###,###.##")
    Next x
    'meNstp
    For x = 0 To 4 Step 1
        txtMiscTotal(x).Text = Format(Val(Replace(txtMiscTotal(x).Text, ",", "")) + Val(Replace(Replace(meNstp(x), ",", ""), "_", "")), "#,###,##0.00")
    Next x

    'Total Cash Basis
    For x = 0 To 4 Step 1
        txtTotalCashBasis(x).Text = Format(Val(Replace(txtOtherFeesTotal(x).Text, ",", "")) + Val(Replace(Replace(txtMiscTotal(x), ",", ""), "_", "")), "#,###,##0.00")
    Next x
    
    
    'Total Downpayment = miscellaneous fee * 20% + other fees
    For x = 0 To 4 Step 1
        txtDownpayment(x).Text = Format(Val(Replace(txtOtherFeesTotal(x).Text, ",", "")) + (Val(Replace(Replace(txtMiscTotal(x), ",", ""), "_", "")) * 0.2), "#,###,##0.00")
    Next x
    
    'Total Installment Basis = Prelim (miscfees * 40%) +
    '                          MidTerm (miscfees * 25%) +
    '                          semifinals (miscfees * 20%) +
    '                          downpayment (miscfees * 20% + otherfees)
    For x = 0 To 4 Step 1
        txtTotalInstallmentBasis(x).Text = Format(Val(Replace(txtMiscTotal(x).Text * ((40 + 25 + 20) / 100), ",", "")) + Val(Replace(Replace(txtDownpayment(x).Text, ",", ""), "_", "")), "#,###,##0.00")
    Next x
    
End Sub


'Focus & Tab emulation
Private Sub meEntrance_GotFocus(Index As Integer)
    FocusMe (meEntrance(Index))
End Sub
Private Sub meTuitionFee_GotFocus(Index As Integer)
    FocusMe (meTuitionFee(Index))
End Sub
Private Sub meRegistration_GotFocus(Index As Integer)
    FocusMe (meRegistration(Index))
End Sub
Private Sub meLibrary_GotFocus(Index As Integer)
    FocusMe (meLibrary(Index))
End Sub
Private Sub meLaboratory_GotFocus(Index As Integer)
    FocusMe (meLaboratory(Index))
End Sub
Private Sub meAthleticFee_GotFocus(Index As Integer)
    FocusMe (meAthleticFee(Index))
End Sub
Private Sub meGuidanceAndCounselor_GotFocus(Index As Integer)
    FocusMe (meGuidanceAndCounselor(Index))
End Sub
Private Sub meRle_GotFocus(Index As Integer)
    FocusMe (meRle(Index))
End Sub
Private Sub meAffiliation_GotFocus(Index As Integer)
    FocusMe (meAffiliation(Index))
End Sub
Private Sub meNursingAudit_GotFocus(Index As Integer)
    FocusMe (meNursingAudit(Index))
End Sub
Private Sub meMarineLaboratory_GotFocus(Index As Integer)
    FocusMe (meMarineLaboratory(Index))
End Sub
Private Sub meSpeechLab_GotFocus(Index As Integer)
    FocusMe (meSpeechLab(Index))
End Sub
Private Sub meHrmLab_GotFocus(Index As Integer)
    FocusMe (meHrmLab(Index))
End Sub
Private Sub meOjt_GotFocus(Index As Integer)
    FocusMe (meOjt(Index))
End Sub
Private Sub meRta_GotFocus(Index As Integer)
    FocusMe (meRta(Index))
End Sub
Private Sub meHon_GotFocus(Index As Integer)
    FocusMe (meRta(Index))
End Sub
Private Sub meMta_GotFocus(Index As Integer)
    FocusMe (meMta(Index))
End Sub
Private Sub meIdNamePlate_GotFocus(Index As Integer)
    FocusMe (meIdNamePlate(Index))
End Sub
Private Sub meSdf_GotFocus(Index As Integer)
    FocusMe (meSdf(Index))
End Sub
Private Sub mePowerFee_GotFocus(Index As Integer)
    FocusMe (mePowerFee(Index))
End Sub
Private Sub meInternet_GotFocus(Index As Integer)
    FocusMe (meInternet(Index))
End Sub
Private Sub meInternship_GotFocus(Index As Integer)
    FocusMe (meInternship(Index))
End Sub
Private Sub meWaiver_GotFocus(Index As Integer)
    FocusMe (meWaiver(Index))
End Sub
Private Sub meNstp_GotFocus(Index As Integer)
    FocusMe (meNstp(Index))
End Sub


'change/keypress
Private Sub meEntrance_Change(Index As Integer)
    ComputeTotal
End Sub
Private Sub meEntrance_KeyPress(Index As Integer, KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub
Private Sub meTuitionFee_Change(Index As Integer)
    ComputeTotal
End Sub
Private Sub meTuitionFee_KeyPress(Index As Integer, KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub
Private Sub meRegistration_Change(Index As Integer)
    ComputeTotal
End Sub
Private Sub meRegistration_KeyPress(Index As Integer, KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub
Private Sub meLibrary_Change(Index As Integer)
    ComputeTotal
End Sub
Private Sub meLibrary_KeyPress(Index As Integer, KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub
Private Sub meLaboratory_Change(Index As Integer)
    ComputeTotal
End Sub
Private Sub meLaboratory_KeyPress(Index As Integer, KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub
Private Sub meAthleticFee_Change(Index As Integer)
    ComputeTotal
End Sub
Private Sub meAthleticFee_KeyPress(Index As Integer, KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub
Private Sub meGuidanceAndCounselor_Change(Index As Integer)
    ComputeTotal
End Sub
Private Sub meGuidanceAndCounselor_KeyPress(Index As Integer, KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub
Private Sub meRle_Change(Index As Integer)
    ComputeTotal
End Sub
Private Sub meRle_KeyPress(Index As Integer, KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub
Private Sub meAffiliation_Change(Index As Integer)
    ComputeTotal
End Sub
Private Sub meAffiliation_KeyPress(Index As Integer, KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub
Private Sub meNursingAudit_Change(Index As Integer)
    ComputeTotal
End Sub
Private Sub meNursingAudit_KeyPress(Index As Integer, KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub
Private Sub meMarineLaboratory_Change(Index As Integer)
    ComputeTotal
End Sub
Private Sub meMarineLaboratory_KeyPress(Index As Integer, KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub
Private Sub meSpeechLab_Change(Index As Integer)
    ComputeTotal
End Sub
Private Sub meSpeechLab_KeyPress(Index As Integer, KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub
Private Sub meHrmLab_Change(Index As Integer)
    ComputeTotal
End Sub
Private Sub meHrmLab_KeyPress(Index As Integer, KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub
Private Sub meOjt_Change(Index As Integer)
    ComputeTotal
End Sub
Private Sub meOjt_KeyPress(Index As Integer, KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub
Private Sub meRta_Change(Index As Integer)
    ComputeTotal
End Sub
Private Sub meRta_KeyPress(Index As Integer, KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub
Private Sub meHon_Change(Index As Integer)
    ComputeTotal
End Sub
Private Sub meHon_KeyPress(Index As Integer, KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub
Private Sub meMta_Change(Index As Integer)
    ComputeTotal
End Sub
Private Sub meMta_KeyPress(Index As Integer, KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub
Private Sub meIdNamePlate_Change(Index As Integer)
    ComputeTotal
End Sub
Private Sub meIdNamePlate_KeyPress(Index As Integer, KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub
Private Sub meSdf_Change(Index As Integer)
    ComputeTotal
End Sub
Private Sub meSdf_KeyPress(Index As Integer, KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub
Private Sub mePowerFee_Change(Index As Integer)
    ComputeTotal
End Sub
Private Sub mePowerFee_KeyPress(Index As Integer, KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub
Private Sub meInternet_Change(Index As Integer)
    ComputeTotal
End Sub
Private Sub meInternet_KeyPress(Index As Integer, KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub
Private Sub meInternship_Change(Index As Integer)
    ComputeTotal
End Sub
Private Sub meInternship_KeyPress(Index As Integer, KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub
Private Sub meWaiver_Change(Index As Integer)
    ComputeTotal
End Sub
Private Sub meWaiver_KeyPress(Index As Integer, KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub
Private Sub meNstp_Change(Index As Integer)
    ComputeTotal
End Sub
Private Sub meNstp_KeyPress(Index As Integer, KeyAscii As Integer)
    EmulateEnter (KeyAscii)
End Sub


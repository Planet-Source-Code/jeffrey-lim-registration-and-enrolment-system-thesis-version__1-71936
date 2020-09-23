Attribute VB_Name = "modPublicVar"
Option Explicit

Public Enc                              As New clsBlowfish

Public strSql                           As String
Global SchoolInformation                As SchoolInfo
Global ReportInformation                As ReportInfo
Global FeesBreakdown                    As FeesBrkdwn


Public User                             As UserInfo 'UserId, UserName, UserType

Global FormOpacity                      As Integer

'--

Global end_app                          As Boolean

'General connection
Public adoConnection As ADODB.Connection    'declare adodb connection variable as public
Public adoCommand As ADODB.Command          'declare adodb command variable as public
Public dbServer As String                   '"192.168.1.80 or (local)"  ' This is the Host computer on a network
                                            ' You may change it into localhost if you are running on a server.
                                            ' or if you are on a network use the computer name or
                                            ' IP address where MS SQL Server resides.
Public dbUser As String                     'default MSSQL username (System Account)
Public dbPassword As String                 'default MSSQL password (blank password)
Public dbName As String                     'database name

'For student
Global frm_stud_show            As Boolean

Global rs_stud                  As New ADODB.Recordset
'For level
Global rs_level                 As New ADODB.Recordset
'For School Year
Global rs_sec                   As New ADODB.Recordset




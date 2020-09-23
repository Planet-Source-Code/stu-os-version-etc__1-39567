VERSION 5.00
Object = "{2BD531E5-B5CD-4EE1-997F-1D96891863EA}#2.0#0"; "dtsystemmonitor.ocx"
Begin VB.Form frmSysInfo 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Information"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   Icon            =   "frmSysInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Index           =   0
      Left            =   360
      Picture         =   "frmSysInfo.frx":030A
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   60
      Width           =   480
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00FF8080&
         Cancel          =   -1  'True
         Caption         =   "Exit"
         DisabledPicture =   "frmSysInfo.frx":074C
         Height          =   375
         Left            =   4320
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   5820
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton cmdSysInfo 
         BackColor       =   &H00FF8080&
         Caption         =   "Built in Windows Info ..."
         Height          =   375
         Left            =   180
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   5820
         Width           =   2055
      End
      Begin SystemMonitor.dtSystemMonitor dtSystemMonitor1 
         Height          =   435
         Left            =   4680
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   300
         Visible         =   0   'False
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   767
         HDDriveLetter   =   "C:\"
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   5340
         Top             =   300
      End
      Begin VB.Label lblSysMonitor 
         BackColor       =   &H00FFC0C0&
         Caption         =   "UNKNOWN"
         Height          =   285
         Index           =   2
         Left            =   2640
         TabIndex        =   27
         Top             =   4920
         Width           =   2145
      End
      Begin VB.Label lblSysMonitor 
         BackColor       =   &H00FFC0C0&
         Caption         =   "UNKNOWN"
         Height          =   285
         Index           =   1
         Left            =   2640
         TabIndex        =   26
         Top             =   4560
         Width           =   2145
      End
      Begin VB.Label lblSysMonitorLabel 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Free Physical Memory :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   4920
         Width           =   2295
      End
      Begin VB.Label lblSysMonitorLabel 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Physical Memory (RAM) :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   4560
         Width           =   2175
      End
      Begin VB.Label lblSysMonitor 
         BackColor       =   &H00FFC0C0&
         Caption         =   "UNKNOWN"
         Height          =   285
         Index           =   10
         Left            =   2640
         TabIndex        =   23
         Top             =   4200
         Width           =   2385
      End
      Begin VB.Label lblSysMonitorLabel 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Available HD Space :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   120
         TabIndex        =   22
         Top             =   4200
         Width           =   2295
      End
      Begin VB.Label lblSysMonitor 
         BackColor       =   &H00FFC0C0&
         Caption         =   "UNKNOWN"
         Height          =   285
         Index           =   9
         Left            =   2640
         TabIndex        =   21
         Top             =   3840
         Width           =   2505
      End
      Begin VB.Label lblSysMonitorLabel 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Total HD Size :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   120
         TabIndex        =   20
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   5880
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lblCRTime 
         BackColor       =   &H00FFC0C0&
         Caption         =   "lblCRTime"
         Height          =   285
         Left            =   2640
         TabIndex        =   18
         Top             =   2040
         Width           =   3255
      End
      Begin VB.Label lblCRT 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Computer Running Time :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   2385
      End
      Begin VB.Label lblSysTime 
         BackColor       =   &H00FFC0C0&
         Caption         =   "lblSysTime"
         Height          =   285
         Left            =   2640
         TabIndex        =   16
         Top             =   2400
         Width           =   3225
      End
      Begin VB.Label lblSys 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Computer Time and Date :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   2400
         Width           =   2385
      End
      Begin VB.Label lblComp 
         BackColor       =   &H00FFC0C0&
         Caption         =   "lblComp"
         Height          =   285
         Left            =   2640
         TabIndex        =   14
         Top             =   1680
         Width           =   3225
      End
      Begin VB.Label lblCN 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Computer Name :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   2085
      End
      Begin VB.Label lblUser 
         BackColor       =   &H00FFC0C0&
         Caption         =   "lblUser"
         Height          =   285
         Left            =   2640
         TabIndex        =   12
         Top             =   1320
         Width           =   3300
      End
      Begin VB.Label lblUN 
         BackColor       =   &H00FFC0C0&
         Caption         =   "User Name :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   2085
      End
      Begin VB.Label lblVersion 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Prog Version"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label lblBuild 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Build"
         Height          =   285
         Left            =   2640
         TabIndex        =   9
         Top             =   3480
         Width           =   2775
      End
      Begin VB.Label lblMin 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Min"
         Height          =   285
         Left            =   4080
         TabIndex        =   8
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label lblMaj 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Maj"
         Height          =   285
         Left            =   2640
         TabIndex        =   7
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label lblName 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Name"
         Height          =   285
         Left            =   2640
         TabIndex        =   6
         Top             =   2760
         Width           =   3255
      End
      Begin VB.Label lblOSBuild 
         BackColor       =   &H00FFC0C0&
         Caption         =   "OS Build Version :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label lblOSMaj 
         BackColor       =   &H00FFC0C0&
         Caption         =   "OS Version :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label lblOSN 
         BackColor       =   &H00FFC0C0&
         Caption         =   "OS Name :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label lblSound 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sound:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   5280
         Width           =   2175
      End
      Begin VB.Label lblCard 
         BackColor       =   &H00FFC0C0&
         Caption         =   "SOUND CARD......................"
         Height          =   525
         Left            =   2640
         TabIndex        =   1
         Top             =   5280
         Width           =   3255
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmSysInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* COPYRIGHT SJS TV SERVICES LTD 2002 ALL RIGHTS RESERVED *
' Compiled by Stu Tyler
' sjstv@btinternet.com

Option Explicit

Private Declare Function waveOutGetDevCaps Lib "winmm.dll" Alias "waveOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEOUTCAPS, ByVal uSize As Long) As Long
Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long

Private Const MAXPNAMELEN = 32

Private Type WAVEOUTCAPS
    wMid As Integer
    wPid As Integer
    vDriverVersion As Long
    szPname As String * MAXPNAMELEN
    dwFormats As Long
    wChannels As Integer
    dwSupport As Long
End Type

Private Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" _
    (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
    KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
    KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

'--------------------------------------------------
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

'------------------------------------------
Dim PCName      As String
Dim p           As Long

Dim i           As Integer
Dim x           As WAVEOUTCAPS

Dim Days        As Long
Dim Hours       As Long
Dim Minutes     As Long
Dim Seconds     As Long
Dim Miliseconds As Long

Private Sub cmdSysInfo_Click()

    Call StartSysInfo

End Sub

Private Sub cmdOK_Click()

    Unload Me

End Sub

Private Sub Form_Load()
    'displays the various info
    
    lblName.Caption = GetOSName
    lblMaj.Caption = "Major ... " & GetOSMajorVer
    lblMin.Caption = "Minor ... " & GetOSMinorVer
    lblBuild.Caption = GetOSBuildVer

    '------------------------------------

    waveOutGetDevCaps 0, x, Len(x)

    i = waveOutGetNumDevs()

    If i > 0 Then

        lblCard.Caption = x.szPname
        
'-----------DIFFERENT OPTIONS YOU CAN USE. JUST MAKE A LABEL--------------------
        'lbl1.Caption = "Formats .............. " & x.dwFormats
        'lbl2.Caption = "Support .............. " & x.dwSupport
        'lbl3.Caption = "Driver Version .... " & x.vDriverVersion
        'lbl4.Caption = "Channels .......... " & x.wChannels
        'lbl5.Caption = "Sound Mid ......" & x.wMid
        'lbl6.Caption = "Sound Pid ......" & x.wPid
'------------------------------------------------------------------------------
    Else
        lblCard.Caption = "Can not retrieve Sound info."
      
    End If

    p = NameOfPC(PCName)

        lblUser.Caption = UserName
            lblComp.Caption = PCName

    'gathers prog info
   
        lblVersion.Caption = "Prog Version No: " & App.Major & "." & _
            App.Minor & "." & App.Revision
    
End Sub

Public Sub StartSysInfo()
    'standard VB Sysinfo stuff

    Dim rc As Long
    Dim SysInfoPath As String

    On Error GoTo SysInfoErr
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
            ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
        ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, 2)
    
    Exit Sub

SysInfoErr:
    MsgBox "Advanced System Information Is Unavailable On This Computer", vbOKOnly

End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable

    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                               ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
        KeyValType, tmpVal, KeyValSize)                     ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win98 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    
        Case REG_SZ                                             ' String Registry Key Data Type
            KeyVal = tmpVal                                     ' Copy String Value
    
        Case REG_DWORD                                          ' Double Word Registry Key Data Type
            For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
                KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
            Next

            KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:          ' Cleanup After An Error Has Occured...
    
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    
End Function

Private Sub Timer1_Timer()

    lblSysTime.Caption = Time & "        " & Format(Date, "dddd, mmm d yyyy")
    lblCRTime.Caption = FormatCount(GetTickCount, DaysHoursMinutesSecondsMilliseconds)

End Sub

Public Property Get UserName() As String

Dim sBuffer As String
Dim lSize As Long
 
    sBuffer = Space$(256)
    lSize = Len(sBuffer)
    
    If GetUserName(sBuffer, lSize) = 0 Then
        Err.Raise vbObjectError + 1, , "Can not retrieve User Name."
    Else
        UserName = Left$(sBuffer, lSize)
    End If
       
End Property

Public Property Get ThreadID() As Variant

    ThreadID = GetCurrentThreadId

End Property

Public Property Get ProcessID() As Variant

    ProcessID = GetCurrentProcessId

End Property

Public Function NameOfPC(MachineName As String) As Long
 
Dim NameSize As Long
Dim x As Long
 
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    x = GetComputerName(MachineName, NameSize)
    
    If x = 0 Then Err.Raise vbObjectError + 2, , "Can not retrieve Computer Name."
 
End Function
 
Private Sub dtSystemMonitor1_Update(ByVal CPULoadPercent As Long, ByVal MemoryLoadPercent As Long, ByVal PhysicalMemoryTotal As Long, ByVal PhysicalMemoryAvailable As Long, ByVal PhysicalMemoryAvailablePercent As Single, ByVal PageFileTotal As Long, ByVal PageFileAvailable As Long, ByVal PageFileAvailablePercent As Single, ByVal VirtualMemoryTotal As Long, ByVal VirtualMemoryAvailable As Long, ByVal VirtualMemoryAvailablePercent As Single, ByVal HDTotalBytes As Currency, ByVal HDTotalFreeBytes As Currency, ByVal HDAvailableFreeBytes As Currency, ByVal HDTotalBytesUsed As Currency, ByVal HDAvailablePercent As Single)

    With dtSystemMonitor1
    
'---------------MORE OPTIONS YOU CAN USE. JUST ADD THE LABEL--------------------

        'lblSysMonitor(0).Caption = Format$(CPULoadPercent, "##0") & " %"
        'lblSysMonitor(7).Caption = Format$(MemoryLoadPercent, "##0") & " %"
        lblSysMonitor(1).Caption = .FormatFilesize(PhysicalMemoryTotal)
        lblSysMonitor(2).Caption = .FormatFilesize(PhysicalMemoryAvailable)
        
        'lblSysMonitor(3).Caption = .FormatFilesize(PageFileTotal)
        'lblSysMonitor(4).Caption = .FormatFilesize(PageFileAvailable)
        
        'lblSysMonitor(5).Caption = .FormatFilesize(VirtualMemoryTotal)
        'lblSysMonitor(6).Caption = .FormatFilesize(VirtualMemoryAvailable)
    
        'lblSysMonitor(8).Caption = .FormatFilesize(HDTotalFreeBytes)
        lblSysMonitor(9).Caption = .FormatFilesize(HDTotalBytes)
        lblSysMonitor(10).Caption = .FormatFilesize(HDAvailableFreeBytes)
        'lblSysMonitor(11).Caption = .FormatFilesize(HDTotalBytesUsed)
        'lblSysMonitor(12).Caption = Format$(HDAvailablePercent, "##0.0") & " %"
        
    End With
    
End Sub

Private Function FormatCount(Count As Long, Optional FormatType As TimeFormatType = 0) As String

    Miliseconds = Count Mod 1000
    Count = Count \ 1000
    Days = Count \ (24& * 3600&)
    If Days > 0 Then Count = Count - (24& * 3600& * Days)
    Hours = Count \ 3600&
    If Hours > 0 Then Count = Count - (3600& * Hours)
    Minutes = Count \ 60
    Seconds = Count Mod 60

    Select Case FormatType
        Case 0

            FormatCount = Days & " days, " & Hours & " hours, " & _
                Minutes & " minutes, " & Seconds & " seconds "
           
    End Select
    
End Function


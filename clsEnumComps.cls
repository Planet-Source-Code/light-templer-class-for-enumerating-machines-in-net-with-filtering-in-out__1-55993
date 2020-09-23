VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEnumComps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'
'   clsEnumComps.cls
'

'   9/3/2004
'   LightTempler

'   V. 1.00

Option Explicit


' *******************************
' *            EVENTS           *
' *******************************
Public Event CompFound(sCompName As String)
Public Event Error(sError As String)


' *******************************
' *            ENUMS            *
' *******************************
Public Enum enCompType
    EC_CT_AllMachines = 1
    EC_CT_AllWindowsComps = 2
    EC_CT_TerminalServer = 3
    EC_CT_Unix = 4
    EC_CT_PrintServer = 5
    EC_CT_DomainController = 6
    EC_CT_SQLServer = 7
    EC_CT_TimeServer = 8
    EC_CT_DFS = 9
End Enum


' *******************************
' *           TYPES             *
' *******************************
Private Type tpSERVER_INFO_API              ' Level 101
    PlatformId  As Long
    servername  As Long
    VerMajor    As Long
    VerMinor    As Long
    Type        As Long
    comment     As Long
End Type

Private Type tpServerInfoFull
    PlatformId  As Long
    servername  As String
    Type        As Long
    VerMajor    As Long
    VerMinor    As Long
    comment     As String
    Platform    As String
    ServerType  As Long
    LanGroup    As String
    LanRoot     As String
End Type



' *************************************
' *            CONSTANTS              *
' *************************************

' Const list from http://vbnet.mvps.org
Private Const SV_TYPE_WORKSTATION         As Long = &H1
Private Const SV_TYPE_SERVER              As Long = &H2
Private Const SV_TYPE_SQLSERVER           As Long = &H4
Private Const SV_TYPE_DOMAIN_CTRL         As Long = &H8
Private Const SV_TYPE_DOMAIN_BAKCTRL      As Long = &H10
Private Const SV_TYPE_TIME_SOURCE         As Long = &H20
Private Const SV_TYPE_AFP                 As Long = &H40
Private Const SV_TYPE_NOVELL              As Long = &H80
Private Const SV_TYPE_DOMAIN_MEMBER       As Long = &H100
Private Const SV_TYPE_PRINTQ_SERVER       As Long = &H200
Private Const SV_TYPE_DIALIN_SERVER       As Long = &H400
Private Const SV_TYPE_XENIX_SERVER        As Long = &H800
Private Const SV_TYPE_SERVER_UNIX         As Long = SV_TYPE_XENIX_SERVER
Private Const SV_TYPE_NT                  As Long = &H1000
Private Const SV_TYPE_WFW                 As Long = &H2000
Private Const SV_TYPE_SERVER_MFPN         As Long = &H4000
Private Const SV_TYPE_SERVER_NT           As Long = &H8000
Private Const SV_TYPE_POTENTIAL_BROWSER   As Long = &H10000
Private Const SV_TYPE_BACKUP_BROWSER      As Long = &H20000
Private Const SV_TYPE_MASTER_BROWSER      As Long = &H40000
Private Const SV_TYPE_DOMAIN_MASTER       As Long = &H80000
Private Const SV_TYPE_SERVER_OSF          As Long = &H100000
Private Const SV_TYPE_SERVER_VMS          As Long = &H200000
Private Const SV_TYPE_WINDOWS             As Long = &H400000            ' Windows95 and above
Private Const SV_TYPE_DFS                 As Long = &H800000            ' Root of a DFS tree
Private Const SV_TYPE_CLUSTER_NT          As Long = &H1000000           ' NT Cluster
Private Const SV_TYPE_TERMINALSERVER      As Long = &H2000000           ' Terminal Server
Private Const SV_TYPE_DCE                 As Long = &H10000000          ' IBM DSS
Private Const SV_TYPE_ALTERNATE_XPORT     As Long = &H20000000          ' rtn alternate transport
Private Const SV_TYPE_LOCAL_LIST_ONLY     As Long = &H40000000          ' rtn local only
Private Const SV_TYPE_DOMAIN_ENUM         As Long = &H80000000
Private Const SV_TYPE_ALL                 As Long = &HFFFFFFFF

Private Const NERR_Success          As Long = 0&
Private Const NERR_MoreData         As Long = 234&



' *************************************
' *        API DEFINITIONS            *
' *************************************
Private Declare Function API_NetServerEnum Lib "NETAPI32" Alias "NetServerEnum" _
        (lpServer As Any, _
         ByVal lLevel As Long, _
         vBuffer As Any, _
         lPreferedMaxLen As Long, _
         lEntriesRead As Long, _
         lTotalEntries As Long, _
         ByVal lServerType As Long, _
         ByVal sDomain$, _
         vResume As Any) As Long

Private Declare Sub API_CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" _
        (pTo As Any, _
         uFrom As Any, _
         ByVal lSize As Long)

Private Declare Function API_lstrlenW Lib "KERNEL32" Alias "lstrlenW" _
        (ByVal PointerToString As Long) As Long


Private Declare Function API_NetAPIBufferFree Lib "NETAPI32" Alias "NetApiBufferFree" _
        (ByVal lpBuffer As Long) As Long



' *************************************
' *            PRIVATES               *
' *************************************
Private mvarFilterIn As String 'local copy
Private mvarFilterOut As String 'local copy
'
'
'

' *************************************
' *         PUBLIC FUNCTIONS          *
' *************************************
Public Function EnumComps(ByVal sDomain As String, ByVal CompType As enCompType) As Long

    ' Sorry, I don't know the original source of this - its the usual way fo enumerating
    ' in Windows API - nothing special.

    Dim tServerInfo         As tpSERVER_INFO_API
    Dim ServerInfo          As tpServerInfoFull
    Dim nRet                As Long
    Dim i                   As Long
    Dim lRetCode            As Long
    Dim lServerInfo         As Long
    Dim lServerInfoPtr      As Long
    Dim lPreferedMaxLen     As Long
    Dim lEntriesRead        As Long
    Dim lTotalEntries       As Long
    Dim yServer()           As Byte
    Dim lServerType         As Long
    Dim sCompName           As String
    Dim lFilterOutLength    As Long
    Dim vResume             As Variant
    
    On Local Error GoTo error_handler


    If sDomain = "" Or CompType = 0 Then
        RaiseEvent Error("Wrong Parameter")
        
        Exit Function
    End If

        
    If mvarFilterIn = "" Then
        mvarFilterIn = "*"
    End If
    mvarFilterIn = UCase$(Trim$(mvarFilterIn))
    mvarFilterOut = UCase$(Trim$(mvarFilterOut))
    
    ' Add here what you need from const list / enum above
    lServerType = Choose(CompType, SV_TYPE_ALL, _
                                    SV_TYPE_NT, _
                                    SV_TYPE_TERMINALSERVER, _
                                    SV_TYPE_SERVER_UNIX, _
                                    SV_TYPE_PRINTQ_SERVER, _
                                    SV_TYPE_DOMAIN_CTRL, _
                                    SV_TYPE_SQLSERVER, _
                                    SV_TYPE_TIME_SOURCE, _
                                    SV_TYPE_DFS)
        
        
    yServer = "" + vbNullChar
    lPreferedMaxLen = 65536
    sDomain = StrConv(sDomain, vbUnicode)
    lFilterOutLength = Len(mvarFilterOut)
    mvarFilterOut = UCase$(mvarFilterOut)
    
    ' Until there is no more data, get list by calling API 'NetServerEnum()'
    nRet = NERR_MoreData
    Do While (nRet = NERR_MoreData)
        
        nRet = API_NetServerEnum(yServer(0), _
                101, _
                lServerInfo, _
                lPreferedMaxLen, _
                lEntriesRead, _
                lTotalEntries, _
                lServerType, _
                sDomain, _
                vResume)
        
        If (nRet <> NERR_Success And nRet <> NERR_MoreData) Then
            
            Exit Do
        End If
        
        ' Get comp names names from list
        i = 1
        lServerInfoPtr = lServerInfo
        Do While i <= lTotalEntries
            API_CopyMemory tServerInfo, ByVal lServerInfoPtr, Len(tServerInfo)
            sCompName = UCase$(PointerToStringW(tServerInfo.servername))
            If sCompName Like mvarFilterIn Then
                If lFilterOutLength Then
                    If Left$(sCompName, lFilterOutLength) <> mvarFilterOut Then
                        RaiseEvent CompFound(sCompName)
                        EnumComps = EnumComps + 1
                    End If
                Else
                    RaiseEvent CompFound(sCompName)
                    EnumComps = EnumComps + 1
                End If
            End If
            i = i + 1
            lServerInfoPtr = lServerInfoPtr + Len(tServerInfo)
        Loop
        lRetCode = API_NetAPIBufferFree(lServerInfo)
    Loop


    Exit Function


error_handler:

    RaiseEvent Error("[" + Err.Description + "] in procedure 'EnumComps()' at class 'clsEnumComps'")
    EnumComps = -1 ' Error

End Function



' *************************************
' *         PRIVATE FUNCTIONS         *
' *************************************
Private Function PointerToStringW(ByVal lpStringW As Long) As String
    
    Dim Buffer()    As Byte
    Dim nLen        As Long
    
    If lpStringW Then
        nLen = API_lstrlenW(lpStringW) * 2
        
        If nLen Then
            ReDim Buffer(0 To (nLen - 1)) As Byte
            API_CopyMemory Buffer(0), ByVal lpStringW, nLen
            PointerToStringW = Buffer
        End If
        
    End If
    
End Function



' *************************************
' *           PROPERTIES              *
' *************************************
Public Property Let FilterOut(ByVal vData As String)
    
    mvarFilterOut = vData

End Property

Public Property Get FilterOut() As String
    
    FilterOut = mvarFilterOut

End Property


Public Property Let FilterIn(ByVal vData As String)
    
    mvarFilterIn = vData

End Property

Public Property Get FilterIn() As String
    
    FilterIn = mvarFilterIn

End Property


' *#*

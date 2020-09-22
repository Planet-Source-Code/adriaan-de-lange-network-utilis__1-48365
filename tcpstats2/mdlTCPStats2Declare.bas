Attribute VB_Name = "mdlTCPStats2Declare"
Option Explicit

Public Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Public Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired As Integer, lpWSAData As WSADATA) As Long
Public Declare Function WSACleanup Lib "wsock32.dll" () As Long
Public Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
Public Declare Function gethostbyaddr Lib "wsock32.dll" (addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long
Public Declare Function gethostname Lib "wsock32.dll" (ByVal hostname$, ByVal HostLen As Long) As Long
Public Declare Function gethostbyname Lib "wsock32.dll" (ByVal hostname$) As Long
Public Declare Function NetMessageBufferSend Lib "netapi32" (ByVal servername As String, ByVal msgname As String, ByVal fromname As String, ByVal msgbuf As String, ByRef msgbuflen As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function NetServerEnum Lib "netapi32" (ByVal servername As Long, ByVal level As Long, buf As Any, ByVal prefmaxlen As Long, entriesread As Long, totalentries As Long, ByVal servertype As Long, ByVal domain As Long, resume_handle As Long) As Long
Public Declare Function NetApiBufferFree Lib "netapi32" (ByVal Buffer As Long) As Long
Public Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Public Declare Function GetTcpTable Lib "iphlpapi.dll" (ByRef pTcpTable As Any, ByRef pdwSize As Long, ByVal bOrder As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Public Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Public Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Long, ByVal RequestOptions As Long, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal Timeout As Long) As Long
Public Declare Function SendARP Lib "iphlpapi.dll" (ByVal DestIP As Long, ByVal SrcIP As Long, pMacAddr As Long, PhyAddrLen As Long) As Long
Public Declare Function NetUserGetInfo Lib "netapi32" (lpServer As Byte, username As Byte, ByVal level As Long, lpBuffer As Long) As Long
Public Declare Function NetUserEnum Lib "netapi32" (servername As Byte, ByVal level As Long, ByVal filter As Long, buff As Long, ByVal buffsize As Long, entriesread As Long, totalentries As Long, resumehandle As Long) As Long
Public Declare Function GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function StrLen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long

Public Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)

Attribute VB_Name = "mdlTCPStats2Type"
Option Explicit

Public Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

Public Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type

Public Type OSVERSIONINFO
  OSVSize         As Long
  dwVerMajor      As Long
  dwVerMinor      As Long
  dwBuildNumber   As Long
  PlatformID      As Long
  szCSDVersion    As String * 128
End Type

Public Type SERVER_INFO_100
  sv100_platform_id As Long
  sv100_name As Long
End Type

Public Type MIB_TCPROW
    dwState As Long
    dwLocalAddr As Long
    dwLocalPort As Long
    dwRemoteAddr As Long
    dwRemotePort As Long
End Type

Private Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type

Public Type ICMP_ECHO_REPLY
    Address         As Long
    status          As Long
    RoundTripTime   As Long
    DataSize        As Long 'formerly integer
   'Reserved        As Integer
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type

Public Type NetMessageData
   sServerName As String
   sSendTo As String
   sSendFrom As String
   sMessage As String
End Type

Public Type USER_INFO_10
   usr10_name          As Long
   usr10_comment       As Long
   usr10_usr_comment   As Long
   usr10_full_name     As Long
End Type

Public Type USER_INFO
   name          As String
   full_name     As String
   comment       As String
   usr_comment   As String
End Type

Attribute VB_Name = "mdlTCPStats2Functions"
Option Explicit

Public Function hibyte(ByVal wParam As Integer)
    hibyte = wParam \ &H100 And &HFF&
End Function

Public Function lobyte(ByVal wParam As Integer)
    lobyte = wParam And &HFF&
End Function

Public Function LocalHostName() As String
Dim hostname As String * 256

    If gethostname(hostname, 256) = SOCKET_ERROR Then
        LocalHostName = "<Error>"
    Else
        LocalHostName = Trim$(hostname)
    End If
End Function

Public Function GetHostByIP(strIP As String) As String
    Dim apiError As Long
    If Len(strIP) < 1 Then Exit Function 'Must contain text
    
    Dim Host As HOSTENT 'Cannot use HOSTENT
    Dim lngIP As Long
    Dim strHost As String * 255
    Dim tmpString As String

    lngIP = inet_addr(strIP & Chr(0))
    
    apiError = gethostbyaddr(lngIP, Len(lngIP), 2)
    If apiError = 0 Then
        GetHostByIP = "Error Receiving HostName..."
        Exit Function
    End If
    
    'Copy mem
    RtlMoveMemory Host, apiError, Len(Host)
    RtlMoveMemory ByVal strHost, Host.hName, 255

    tmpString = strHost 'I think you can use strHost
    
    'Pull from beginning to null
    If InStr(tmpString, Chr(0)) <> 0 Then
        tmpString = Left(tmpString, InStr(tmpString, Chr(0)) - 1)
    End If
    
    tmpString = Trim(tmpString)
    
    GetHostByIP = tmpString 'Send back out
End Function

Public Function GetIPByHost(strHost As String) As String
    Dim apiError As Long
    If Len(strHost) < 1 Then Exit Function 'Must contain text
    
    Dim Host As HOSTENT   'Cannot use HOSTENT
    Dim lngHostIp As Long
    Dim strIP As String
    Dim tmpIP() As Byte '1/4 ip = byte
    Dim tmpInt As Integer
    
    apiError = gethostbyname(strHost & Chr(0))
    If apiError = 0 Then
        GetIPByHost = "Error Receiving IP Address..."
        Exit Function
    End If

    'Copy mem
    RtlMoveMemory Host, apiError, LenB(Host)
    RtlMoveMemory lngHostIp, Host.hAddrList, 4  'Copy 4 parts of ip

    ReDim tmpIP(1 To Host.hLength)  'Resize
    RtlMoveMemory tmpIP(1), lngHostIp, Host.hLength  'Copy mem

    For tmpInt = 1 To Host.hLength  'Cyle through all parts
        strIP = strIP & tmpIP(tmpInt) & "." 'Add . in between
    Next
    strIP = Mid(strIP, 1, Len(strIP) - 1) 'Remove extra .

    GetIPByHost = strIP 'Send back out
End Function

Public Function Valid_IP(IP As String) As Boolean
    Dim i As Integer
    Dim dot_count As Integer
    Dim test_octet As String
    Dim byte_check
     IP = Trim$(IP)

     ' make sure the IP long enough before
     ' continuing
     If Len(IP) < 8 Then
        Valid_IP = False
        'Show Message
        'MsgBox IP & " is Invalid", , "IP Validator"
        Exit Function
    End If

    i = 1
    dot_count = 0
    For i = 1 To Len(IP)
        If Mid$(IP, i, 1) = "." Then
            ' increment the dot count and
            ' clear the test octet variable
            dot_count = dot_count + 1
            test_octet = ""
            If i = Len(IP) Then
                ' we've ended with a dot
                ' this is not good
                Valid_IP = False
                'Show Message
                'MsgBox IP & " is Invalid", , "IP Validator"
                Exit Function
            End If
        Else
            test_octet = test_octet & Mid$(IP, i, 1)
            On Error Resume Next
            byte_check = CByte(test_octet)
            If (Err) Then
                ' either the value is not numeric
                ' or exceeds the range of the byte
                ' data type.
                Valid_IP = False
                Exit Function
            End If
        End If
    Next i
     ' so far, so good
      ' did we get the correct number of dots?
    If dot_count <> 3 Then
        Valid_IP = False
        Exit Function
    End If
     ' we have a valid IP format!
    Valid_IP = True
        'Show Message
        'MsgBox IP & " is Valid", , "IP Validator"
    
End Function

Public Function GetServersMSGcmb(sDomain As String) As Long

  'lists all servers of the specified type
  'that are visible in a domain.
  
   Dim bufptr          As Long
   Dim dwEntriesread   As Long
   Dim dwTotalentries  As Long
   Dim dwResumehandle  As Long
   Dim se100           As SERVER_INFO_100
   Dim success         As Long
   Dim nStructSize     As Long
   Dim cnt             As Long

   nStructSize = LenB(se100)
   
  'Call passing MAX_PREFERRED_LENGTH to have the
  'API allocate required memory for the return values.
  '
  'The call is enumerating all machines on the
  'network (SV_TYPE_ALL); however, by Or'ing
  'specific bit masks for defined types you can
  'customize the returned data. For example, a
  'value of 0x00000003 combines the bit masks for
  'SV_TYPE_WORKSTATION (0x00000001) and
  'SV_TYPE_SERVER (0x00000002).
  '
  'dwServerName must be Null. The level parameter
  '(100 here) specifies the data structure being
  'used (in this case a SERVER_INFO_100 structure).
  '
  'The domain member is passed as Null, indicating
  'machines on the primary domain are to be retrieved.
  'If you decide to use this member, pass
  'StrPtr("domain name"), not the string itself.
   success = NetServerEnum(0&, _
                           100, _
                           bufptr, _
                           MAX_PREFERRED_LENGTH, _
                           dwEntriesread, _
                           dwTotalentries, _
                           SV_TYPE_ALL, _
                           0&, _
                           dwResumehandle)

  'if all goes well
   If success = NERR_SUCCESS And _
      success <> ERROR_MORE_DATA Then
      
    'loop through the returned data, adding each
    'machine to the list
      For cnt = 0 To dwEntriesread - 1
         
        'get one chunk of data and cast
        'into an SERVER_INFO_100 struct
        'in order to add the name to a list
         CopyMemory se100, ByVal bufptr + (nStructSize * cnt), nStructSize
            
         frmMSGSender.Combo1.AddItem GetPointerToByteStringW(se100.sv100_name)
         frmMSGSender.Combo2.AddItem GetPointerToByteStringW(se100.sv100_name)
         frmMSGSender.Combo3.AddItem GetPointerToByteStringW(se100.sv100_name)
         frmIPResolver.cmbHostName.AddItem GetPointerToByteStringW(se100.sv100_name)
      Next
      
   End If
   
  'clean up regardless of success
   Call NetApiBufferFree(bufptr)
   
  'return entries as sign of success
   GetServersMSGcmb = dwEntriesread

End Function

Public Function GetPointerToByteStringW(ByVal dwData As Long) As String
  
   Dim tmp() As Byte
   Dim tmplen As Long
   
   If dwData <> 0 Then
   
      tmplen = lstrlenW(dwData) * 2
      
      If tmplen <> 0 Then
      
         ReDim tmp(0 To (tmplen - 1)) As Byte
         CopyMemory tmp(0), ByVal dwData, tmplen
         GetPointerToByteStringW = tmp
         
     End If
     
   End If
    
End Function

Public Function GetNetSendMessageStatus(nError As Long) As String
    
   Dim msg As String
   
   Select Case nError
   
     Case NERR_SUCCESS:            msg = "The message was successfully sent"
     Case NERR_NameNotFound:       msg = "Send To not found"
     Case NERR_NetworkError:       msg = "General network error occurred"
     Case NERR_UseNotFound:        msg = "Network connection not found"
     Case ERROR_ACCESS_DENIED:     msg = "Access to computer denied"
     Case ERROR_BAD_NETPATH:       msg = "Sent From server name not found."
     Case ERROR_INVALID_PARAMETER: msg = "Invalid parameter(s) specified."
     Case ERROR_NOT_SUPPORTED:     msg = "Network request not supported."
     Case ERROR_INVALID_NAME:      msg = "Illegal character or malformed name."
     Case Else:                    msg = "Unknown error executing command."
     
   End Select
   
   GetNetSendMessageStatus = msg
   
End Function

Public Function IsWinNT() As Boolean

  'returns True if running WinNT/Win2000/WinXP
   #If Win32 Then
  
      Dim OSV As OSVERSIONINFO
   
      OSV.OSVSize = Len(OSV)
   
      If GetVersionEx(OSV) = 1 Then
   
        'PlatformId contains a value representing the OS.
         IsWinNT = (OSV.PlatformID = VER_PLATFORM_WIN32_NT)
         
      End If

   #End If

End Function

Public Function NetSendMessage(msgData As NetMessageData) As String

   Dim success As Long
   
  'assure that the OS is NT ..
  'NetMessageBufferSend  can not
  'be called on Win9x
   If IsWinNT() Then
      
      With msgData
      
        'if To name omitted return error and exit
         If .sSendTo = "" Then
            
            NetSendMessage = GetNetSendMessageStatus(ERROR_INVALID_PARAMETER)
            Exit Function
            
         Else
       
           'if there is a message
            If Len(.sMessage) Then
   
              'convert the strings to unicode
               .sSendTo = StrConv(.sSendTo, vbUnicode)
               .sMessage = StrConv(.sMessage, vbUnicode)
            
              'Note that the API could be called passing
              'vbNullString as the SendFrom and sServerName
              'strings. This would generate the message on
              'the sending machine.
               If Len(.sServerName) > 0 Then
                     .sServerName = StrConv(.sServerName, vbUnicode)
               Else: .sServerName = vbNullString
               End If
                        
               If Len(.sSendFrom) > 0 Then
                     .sSendFrom = StrConv(.sSendFrom, vbUnicode)
               Else: .sSendFrom = vbNullString
               End If
            
              'change the cursor and show. Control won't return
              'until the call has completed.
               Screen.MousePointer = vbHourglass
           
               success = NetMessageBufferSend(.sServerName, _
                                              .sSendTo, _
                                              .sSendFrom, _
                                              .sMessage, _
                                              ByVal Len(.sMessage))
           
               Screen.MousePointer = vbNormal
           
               NetSendMessage = GetNetSendMessageStatus(success)
   
            End If 'If Len(.sMessage)
         End If  'If .sSendTo
      End With  'With msgData
   End If  'If IsWinNT
   
End Function

Public Function GetStatusCode(status As Long) As String

   Dim msg As String
   
   Select Case status
      Case IP_SUCCESS:               msg = "ip success"
      Case INADDR_NONE:              msg = "inet_addr: bad IP format"
      Case IP_BUF_TOO_SMALL:         msg = "ip buf too_small"
      Case IP_DEST_NET_UNREACHABLE:  msg = "ip dest net unreachable"
      Case IP_DEST_HOST_UNREACHABLE: msg = "ip dest host unreachable"
      Case IP_DEST_PROT_UNREACHABLE: msg = "ip dest prot unreachable"
      Case IP_DEST_PORT_UNREACHABLE: msg = "ip dest port unreachable"
      Case IP_NO_RESOURCES:          msg = "ip no resources"
      Case IP_BAD_OPTION:            msg = "ip bad option"
      Case IP_HW_ERROR:              msg = "ip hw_error"
      Case IP_PACKET_TOO_BIG:        msg = "ip packet too_big"
      Case IP_REQ_TIMED_OUT:         msg = "ip req timed out"
      Case IP_BAD_REQ:               msg = "ip bad req"
      Case IP_BAD_ROUTE:             msg = "ip bad route"
      Case IP_TTL_EXPIRED_TRANSIT:   msg = "ip ttl expired transit"
      Case IP_TTL_EXPIRED_REASSEM:   msg = "ip ttl expired reassem"
      Case IP_PARAM_PROBLEM:         msg = "ip param_problem"
      Case IP_SOURCE_QUENCH:         msg = "ip source quench"
      Case IP_OPTION_TOO_BIG:        msg = "ip option too_big"
      Case IP_BAD_DESTINATION:       msg = "ip bad destination"
      Case IP_ADDR_DELETED:          msg = "ip addr deleted"
      Case IP_SPEC_MTU_CHANGE:       msg = "ip spec mtu change"
      Case IP_MTU_CHANGE:            msg = "ip mtu_change"
      Case IP_UNLOAD:                msg = "ip unload"
      Case IP_ADDR_ADDED:            msg = "ip addr added"
      Case IP_GENERAL_FAILURE:       msg = "ip general failure"
      Case IP_PENDING:               msg = "ip pending"
      Case PING_TIMEOUT:             msg = "ping timeout"
      Case Else:                     msg = "unknown  msg returned"
   End Select
   
   GetStatusCode = CStr(status) & "   [ " & msg & " ]"
   
End Function

Public Function PingF(sAddress As String, _
                     sDataToSend As String, _
                     ECHO As ICMP_ECHO_REPLY) As Long

  'If Ping succeeds :
  '.RoundTripTime = time in ms for the ping to complete,
  '.Data is the data returned (NULL terminated)
  '.Address is the Ip address that actually replied
  '.DataSize is the size of the string in .Data
  '.Status will be 0
  '
  'If Ping fails .Status will be the error code
   
   Dim hPort As Long
   Dim dwAddress As Long
   
  'convert the address into a long representation
   dwAddress = inet_addr(sAddress)
   
  'if a valid address..
   If dwAddress <> INADDR_NONE Then
   
     'open a port
      hPort = IcmpCreateFile()
      
     'and if successful,
      If hPort Then
      
        'ping it.
         Call IcmpSendEcho(hPort, _
                           dwAddress, _
                           sDataToSend, _
                           Len(sDataToSend), _
                           0, _
                           ECHO, _
                           Len(ECHO), _
                           PING_TIMEOUT)

        'return the status as ping succes and close
         PingF = ECHO.status
         Call IcmpCloseHandle(hPort)
      
      End If
      
   Else:
        'the address format was probably invalid
         PingF = INADDR_NONE
         
   End If
  
End Function

Public Sub SocketsCleanup()
   
   If WSACleanup() <> 0 Then
       MsgBox "Windows Sockets error occurred in Cleanup.", vbExclamation
   End If
    
End Sub

Public Function SocketsInitialize() As Boolean

   Dim WSAD As WSADATA
   
   SocketsInitialize = WSAStartup(WS_VERSION_REQD, WSAD) = IP_SUCCESS
    
End Function

Public Function GetRemoteMACAddress(ByVal sRemoteIP As String, _
                                     sRemoteMacAddress As String) As Boolean

   Dim dwRemoteIP As Long
   Dim pMacAddr As Long
   Dim bpMacAddr() As Byte
   Dim PhyAddrLen As Long
   Dim cnt As Long
   Dim tmp As String
    
  'convert the string IP into
  'an unsigned long value containing
  'a suitable binary representation
  'of the Internet address given
   dwRemoteIP = inet_addr(sRemoteIP)
   
   If dwRemoteIP <> 0 Then
   
     'set PhyAddrLen to 6
      PhyAddrLen = 6
   
     'retrieve the remote MAC address
      If SendARP(dwRemoteIP, 0&, pMacAddr, PhyAddrLen) = NO_ERROR Then
      
         If pMacAddr <> 0 And PhyAddrLen <> 0 Then
      
           'returned value is a long pointer
           'to the mac address, so copy data
           'to a byte array
            ReDim bpMacAddr(0 To PhyAddrLen - 1)
            CopyMemory bpMacAddr(0), pMacAddr, ByVal PhyAddrLen
          
           'loop through array to build string
            For cnt = 0 To PhyAddrLen - 1
               
               If bpMacAddr(cnt) = 0 Then
                  tmp = tmp & "00-"
               Else
                  tmp = tmp & Hex$(bpMacAddr(cnt)) & "-"
               End If
         
            Next
           
           'remove the trailing dash
           'added above and return True
            If Len(tmp) > 0 Then
               sRemoteMacAddress = Left$(tmp, Len(tmp) - 1)
               GetRemoteMACAddress = True
            End If

            Exit Function
         
         Else
            GetRemoteMACAddress = False
         End If
            
      Else
         GetRemoteMACAddress = False
      End If  'SendARP
      
   Else
      GetRemoteMACAddress = False
   End If  'dwRemoteIP
      
End Function

Public Function GetUserEnumInfo(bServername() As Byte)
  
   Dim users() As Long
   Dim buff As Long
   Dim buffsize As Long
   Dim entriesread As Long
   Dim totalentries As Long
   Dim cnt As Integer
   
   buffsize = 255
   
   If NetUserEnum(bServername(0), 0, _
                   FILTER_NORMAL_ACCOUNT, _
                   buff, buffsize, _
                   entriesread, _
                   totalentries, 0&) = ERROR_SUCCESS Then
   
      ReDim users(0 To entriesread - 1) As Long
      CopyMemory users(0), ByVal buff, entriesread * 4
      
      For cnt = 0 To entriesread - 1
        frmUserNames.List1.AddItem GetPointerToByteStringW(users(cnt))
      Next cnt
      
      NetApiBufferFree buff
   
   End If

End Function

Public Function GetComputersName() As String

  'returns the name of the computer
   Dim tmp As String
   
   tmp = Space$(MAX_COMPUTERNAME + 1)
    
   If GetComputerName(tmp, Len(tmp)) <> 0 Then
      GetComputersName = TrimNull(tmp)
   End If
   
End Function

Public Function TrimNull(item As String)

   Dim pos As Integer
   
   pos = InStr(item, Chr$(0))
   
   If pos Then
         TrimNull = Left$(item, pos - 1)
   Else: TrimNull = item
   End If
   
End Function

Public Function GetUserNetworkInfo(bServername() As Byte, bUsername() As Byte) As USER_INFO
   
   Dim usrapi As USER_INFO_10
   Dim buff As Long
   
   If NetUserGetInfo(bServername(0), bUsername(0), 10, buff) = ERROR_SUCCESS Then
      
     'copy the data from buff into the
     'API user_10 structure
      CopyMemory usrapi, ByVal buff, Len(usrapi)
      
     'extract each member and return
     'as members of the UDT
      GetUserNetworkInfo.name = GetPointerToByteStringW(usrapi.usr10_name)
      GetUserNetworkInfo.full_name = GetPointerToByteStringW(usrapi.usr10_full_name)
      GetUserNetworkInfo.comment = GetPointerToByteStringW(usrapi.usr10_comment)
      GetUserNetworkInfo.usr_comment = GetPointerToByteStringW(usrapi.usr10_usr_comment)
   
      NetApiBufferFree buff
   
   End If
   
End Function

Public Function GetServers(sDomain As String) As Long

  'lists all servers of the specified type
  'that are visible in a domain.
  
   Dim bufptr          As Long
   Dim dwEntriesread   As Long
   Dim dwTotalentries  As Long
   Dim dwResumehandle  As Long
   Dim se100           As SERVER_INFO_100
   Dim success         As Long
   Dim nStructSize     As Long
   Dim cnt             As Long

   nStructSize = LenB(se100)
   
  'Call passing MAX_PREFERRED_LENGTH to have the
  'API allocate required memory for the return values.
  '
  'The call is enumerating all machines on the
  'network (SV_TYPE_ALL); however, by Or'ing
  'specific bit masks for defined types you can
  'customize the returned data. For example, a
  'value of 0x00000003 combines the bit masks for
  'SV_TYPE_WORKSTATION (0x00000001) and
  'SV_TYPE_SERVER (0x00000002).
  '
  'dwServerName must be Null. The level parameter
  '(100 here) specifies the data structure being
  'used (in this case a SERVER_INFO_100 structure).
  '
  'The domain member is passed as Null, indicating
  'machines on the primary domain are to be retrieved.
  'If you decide to use this member, pass
  'StrPtr("domain name"), not the string itself.
   success = NetServerEnum(0&, _
                           100, _
                           bufptr, _
                           MAX_PREFERRED_LENGTH, _
                           dwEntriesread, _
                           dwTotalentries, _
                           SV_TYPE_ALL, _
                           0&, _
                           dwResumehandle)

  'if all goes well
   If success = NERR_SUCCESS And _
      success <> ERROR_MORE_DATA Then
      
    'loop through the returned data, adding each
    'machine to the list
      For cnt = 0 To dwEntriesread - 1
         
        'get one chunk of data and cast
        'into an SERVER_INFO_100 struct
        'in order to add the name to a list
         CopyMemory se100, ByVal bufptr + (nStructSize * cnt), nStructSize
            
        frmUserNames.Combo1.AddItem GetPointerToByteStringW(se100.sv100_name)
         
      Next
      
   End If
   
  'clean up regardless of success
   Call NetApiBufferFree(bufptr)
   
  'return entries as sign of success
   GetServers = dwEntriesread

End Function


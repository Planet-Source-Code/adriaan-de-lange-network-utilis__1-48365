Attribute VB_Name = "mdlTCPStats2Sub"
Option Explicit

Public Sub InitializeSockets()
Dim WSAD As WSADATA
Dim iReturn As Integer
Dim sLowByte As String, sHighByte As String, sMsg As String

    iReturn = WSAStartup(WS_VERSION_REQD, WSAD)

    If iReturn <> 0 Then
        frmIPResolver.txtIPAddress.Text = "Winsock.dll is not responding."
        End
    End If

    If lobyte(WSAD.wversion) < WS_VERSION_MAJOR Or (lobyte(WSAD.wversion) = _
        WS_VERSION_MAJOR And hibyte(WSAD.wversion) < WS_VERSION_MINOR) Then

        sHighByte = Trim$(Str$(hibyte(WSAD.wversion)))
        sLowByte = Trim$(Str$(lobyte(WSAD.wversion)))
        sMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
        sMsg = sMsg & " is not supported by winsock.dll "
        frmIPResolver.txtIPAddress = sMsg
        End
    End If

    'iMaxSockets is not used in winsock 2. So the following check is only
    'necessary for winsock 1. If winsock 2 is requested,
    'the following check can be skipped.

    If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
        sMsg = "This application requires a minimum of "
        sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
        frmIPResolver.txtIPAddress.Text = sMsg
        End
    End If

End Sub

Public Sub CleanupSockets()
Dim lReturn As Long

    lReturn = WSACleanup()

    If lReturn <> 0 Then
        frmIPResolver.txtIPAddress.Text = "Socket Error " & Trim$(Str$(lReturn)) & " Occurred In Cleanup."
        End
    End If

End Sub



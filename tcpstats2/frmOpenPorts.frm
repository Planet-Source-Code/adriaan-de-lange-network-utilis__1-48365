VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpenPorts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   407
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   609
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Timer tmrOpenPorts 
      Interval        =   200
      Left            =   4440
      Top             =   5520
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Avoid Local IP's... 0.0.0.0 or 127.0.0.1"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9128
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "ip"
         Text            =   "IP Address"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Port"
         Text            =   "Local Port"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Remote"
         Text            =   "Remote IP"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Remote Port"
         Text            =   "Remote Port"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "State"
         Text            =   "State"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Menu mnuApp 
      Caption         =   "Application"
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmOpenPorts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const ERROR_SUCCESS = 0
Private Current As Long
Private Finished As Boolean

Private Function GetIpFromLong(lngIPAddress As Long) As String
    '
    Dim arrIpParts(3) As Byte
    '
    CopyMemory arrIpParts(0), lngIPAddress, 4
    '
    GetIpFromLong = CStr(arrIpParts(0)) & "." & _
                    CStr(arrIpParts(1)) & "." & _
                    CStr(arrIpParts(2)) & "." & _
                    CStr(arrIpParts(3))
    '
End Function

Private Function GetTcpPortNumber(DWord As Long) As Long
    GetTcpPortNumber = DWord / 256 + (DWord Mod 256) * 256
End Function

Private Function GetState(lngstate As Long) As String
    
    Select Case lngstate
        Case MIB_TCP_STATE_CLOSED: GetState = "CLOSED"
        Case MIB_TCP_STATE_LISTEN: GetState = "LISTEN"
        Case MIB_TCP_STATE_SYN_SENT: GetState = "SYN_SENT"
        Case MIB_TCP_STATE_SYN_RCVD: GetState = "SYN_RCVD"
        Case MIB_TCP_STATE_ESTAB: GetState = "ESTABLISHED"
        Case MIB_TCP_STATE_FIN_WAIT1: GetState = "FIN_WAIT1"
        Case MIB_TCP_STATE_FIN_WAIT2: GetState = "FIN_WAIT2"
        Case MIB_TCP_STATE_CLOSE_WAIT: GetState = "CLOSE_WAIT"
        Case MIB_TCP_STATE_CLOSING: GetState = "CLOSING"
        Case MIB_TCP_STATE_LAST_ACK: GetState = "LAST_ACK"
        Case MIB_TCP_STATE_TIME_WAIT: GetState = "TIME_WAIT"
        Case MIB_TCP_STATE_DELETE_TCB: GetState = "DELETE_TCB"
    End Select
    '
End Function

Private Sub Form_Load()

Me.Show

    'Do While Not Finished
    '    DoEvents
    '    If Finished Then Exit Do
    '
    '    Loop
   
    
End Sub


Private Sub mnuClose_Click()

'Finished = True
Me.Visible = False
Unload Me

End Sub

Private Sub tmrOpenPorts_Timer()
        Do While GetTickCount - Current >= 1000
           
            Current = GetTickCount
            
            Dim arrBuffer() As Byte
            Dim lngSize As Long
            Dim lngRetVal As Long
            Dim lngRows As Long
            Dim i As Long
            Dim TcpTableRow As MIB_TCPROW
            Dim lvItem As ListItem
            
            ListView1.ListItems.Clear
    
            lngSize = 0
    
            'Call the GetTcpTable just to get
            'the buffer size into the lngSize variable
            lngRetVal = GetTcpTable(ByVal 0&, lngSize, 0)
    
            'Prepare the buffer
            ReDim arrBuffer(0 To lngSize - 1) As Byte
    
            'And call the function one more time
            lngRetVal = GetTcpTable(arrBuffer(0), lngSize, 0)
    
                If lngRetVal = ERROR_SUCCESS Then
        
                'The first 4 bytes contain the quantity of the table rows
                'Get that value to the lngRows variable
                CopyMemory lngRows, arrBuffer(0), 4
        
                    For i = 1 To lngRows
            
                        'Copy the table row data to the TcpTableRow structure
                        CopyMemory TcpTableRow, arrBuffer(4 + (i - 1) * _
                            Len(TcpTableRow)), Len(TcpTableRow)
            
                        If Not ((Check1.Value = vbChecked) And _
                            (GetIpFromLong(TcpTableRow.dwLocalAddr) = "0.0.0.0" Or _
                            GetIpFromLong(TcpTableRow.dwLocalAddr) = "127.0.0.1")) Then
                
                            'Add the data to the ListView control
                            With TcpTableRow
                                Set lvItem = ListView1.ListItems.Add(, , _
                                        GetIpFromLong(.dwLocalAddr))
                                lvItem.SubItems(1) = GetTcpPortNumber(.dwLocalPort)
                                lvItem.SubItems(2) = GetIpFromLong(.dwRemoteAddr)
                                lvItem.SubItems(3) = GetTcpPortNumber(.dwRemotePort)
                                lvItem.SubItems(4) = GetState(.dwState)
                            End With
                
                        End If
            
                    Next i
        
                End If
            Loop
            

End Sub

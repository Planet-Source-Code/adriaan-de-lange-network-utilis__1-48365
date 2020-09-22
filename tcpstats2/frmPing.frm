VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPing 
   Caption         =   "Form1"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   255
      Left            =   2640
      TabIndex        =   24
      Top             =   3720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Max             =   100000
      Enabled         =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   12
      Text            =   "96.96.96.12"
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Text            =   "Echo This"
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   10
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   9
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   2
      Left            =   1920
      TabIndex        =   8
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   3
      Left            =   1920
      TabIndex        =   7
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   4
      Left            =   1920
      TabIndex        =   6
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   5
      Left            =   1920
      TabIndex        =   5
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ping && Get Mac Add"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   4080
      Width           =   2535
   End
   Begin VB.TextBox txtRMacA 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox txtCounterV1 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Text            =   "0"
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ping Address"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Send"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Return Status"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Address (dec)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Round Trip Time"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Data Packet Size"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Data Returned"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Data Pointer"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lblRMacA 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Remote Mac Address:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Remote Mac Address:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ping Retries"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   3720
      Width           =   975
   End
End
Attribute VB_Name = "frmPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim TimeNow1, TimeNow2
TimeNow1 = Now()

For bbbb = 0 To txtCounterV1

   Dim ECHO As ICMP_ECHO_REPLY
   Dim pos As Long
   Dim success As Long
   
   If SocketsInitialize() Then
   
     'ping the ip passing the address, text
     'to send, and the ECHO structure.
      success = PingF((Text1.Text), (Text2.Text), ECHO)
      
     'display the results
      Text4(0).Text = GetStatusCode(success)
      Text4(1).Text = ECHO.Address
      Text4(2).Text = ECHO.RoundTripTime & " ms"
      Text4(3).Text = ECHO.DataSize & " bytes"
      
      If Left$(ECHO.Data, 1) <> Chr$(0) Then
         pos = InStr(ECHO.Data, Chr$(0))
         Text4(4).Text = Left$(ECHO.Data, pos - 1)
      End If
   
      Text4(5).Text = ECHO.DataPointer
      
      SocketsCleanup
      
   Else
   
        MsgBox "Windows Sockets for 32 bit Windows " & _
               "environments is not successfully responding."
   
   End If
   

   Dim sRemoteMacAddress As String
   
   If Len(Text1.Text) > 0 Then
   
      If GetRemoteMACAddress(Text1.Text, sRemoteMacAddress) Then
         txtRMacA.Text = sRemoteMacAddress
      Else
         txtRMacA.Text = "(SendARP call failed)"
      End If
      
   End If

Text5 = bbbb
Text5.Refresh
Text4(0).Refresh
Text4(1).Refresh
Text4(2).Refresh
Text4(3).Refresh
Text4(4).Refresh
Text4(5).Refresh

Next bbbb

TimeNow2 = Now()
timeend = TimeNow2 - TimeNow1
Text3 = Format$(timeend, " hh:mm:ss ") & vbNewLine & Format$(TimeNow1, " hh:mm:ss ") & vbNewLine & Format$(TimeNow2, " hh:mm:ss ")

End Sub

Private Sub txtCounterV1_Change()
On Error GoTo EXITSUB
If txtCounterV1.Text > UpDown1.Max Then GoTo EXITSUB
UpDown1.Value = txtCounterV1

Exit Sub

EXITSUB:
If Err.Number = 13 Then txtCounterV1 = 0: Exit Sub
UpDown1.Value = UpDown1.Max

End Sub

Private Sub UpDown1_Change()
txtCounterV1 = UpDown1.Value
End Sub



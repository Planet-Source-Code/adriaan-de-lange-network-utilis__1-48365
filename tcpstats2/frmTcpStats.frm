VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTcpStats 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCP Statistics"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7950
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmTcpStats.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOpenPorts 
      Caption         =   "&Open Ports"
      Height          =   375
      Left            =   6480
      TabIndex        =   9
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdIPResolver 
      Caption         =   "&Resolv IP"
      Height          =   375
      Left            =   6480
      TabIndex        =   8
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdMSGSender 
      Caption         =   "&MSG Sender"
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdUserNames 
      Caption         =   "&User Names"
      Height          =   375
      Left            =   6480
      TabIndex        =   6
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmdPing 
      Caption         =   "&Ping"
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      Top             =   4320
      Width           =   1335
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   6720
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTcpStats.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTcpStats.frx":0466
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTcpStats.frx":0782
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTcpStats.frx":0A9E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7320
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTcpStats.frx":0DBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTcpStats.frx":0F16
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTcpStats.frx":1072
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTcpStats.frx":11CE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parameter description"
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Width           =   7695
      Begin VB.Label lblDesc 
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   7455
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4575
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   8070
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Parameter"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   6480
      Top             =   1080
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "&Hide"
      Height          =   375
      Left            =   6480
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   6840
      Top             =   1080
      Width           =   600
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mPopHide 
         Caption         =   "&Hide"
      End
      Begin VB.Menu mPopRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mPopOpenPorts 
         Caption         =   "&Open Ports"
      End
      Begin VB.Menu mPopIPResolver 
         Caption         =   "&IP Resolver"
      End
      Begin VB.Menu mPopMSGSender 
         Caption         =   "&MSG Sender"
      End
      Begin VB.Menu mPopUserNames 
         Caption         =   "&User Names"
      End
      Begin VB.Menu mPopPing 
         Caption         =   "&Ping"
      End
      Begin VB.Menu mPopExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmTcpStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'user defined type required by Shell_NotifyIcon API call
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'constants required by Shell_NotifyIcon API call:
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up
Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private nid As NOTIFYICONDATA

'user defined type required by GetTcpStatistics API call
Private Type MIB_TCPSTATS
    dwRtoAlgorithm  As Long '// timeout algorithm
    dwRtoMin        As Long '// minimum timeout
    dwRtoMax        As Long '// maximum timeout
    dwMaxConn       As Long '// maximum connections
    dwActiveOpens   As Long '// active opens
    dwPassiveOpens  As Long '// passive opens
    dwAttemptFails  As Long '// failed attempts
    dwEstabResets   As Long '// establised connections reset
    dwCurrEstab     As Long '// established connections
    dwInSegs        As Long '// segments received
    dwOutSegs       As Long '// segment sent
    dwRetransSegs   As Long '// segments retransmitted
    dwInErrs        As Long '// incoming errors
    dwOutRsts       As Long '// outgoing resets
    dwNumConns      As Long '// cumulative connections
End Type

Private Declare Function GetTcpStatistics Lib "iphlpapi.dll" (pStats As MIB_TCPSTATS) As Long

Private Sub cmdExit_Click()
    Unload Me
    End
End Sub

Private Sub cmdHide_Click()
    Me.Hide
    frmping.Hide
End Sub


Private Sub cmdIPResolver_Click()
frmIPResolver.Visible = True
Load frmIPResolver
End Sub

Private Sub cmdMSGSender_Click()
frmMSGSender.Visible = True
Load frmMSGSender
End Sub

Private Sub cmdOpenPorts_Click()
frmOpenPorts.Visible = True
Load frmOpenPorts
End Sub

Private Sub cmdPing_Click()
frmping.Visible = True
Load frmping
End Sub

Private Sub cmdUserNames_Click()
frmUserNames.Visible = True
Load frmUserNames
End Sub

Private Sub Form_Load()
    '
    'Configure the ListView control
    '
    With ListView1.ListItems
        '
        .Add , , "Timeout algorithm"
        .Add , , "Minimum timeout"
        .Add , , "Maximum timeout"
        .Add , , "Maximum connections"
        .Add , , "Active opens"
        .Add , , "Passive opens"
        .Add , , "Failed attempts"
        .Add , , "Establised connections reset"
        .Add , , "Established connections"
        .Add , , "Segments received"
        .Add , , "Segment sent"
        .Add , , "Segments retransmitted"
        .Add , , "Incoming errors"
        .Add , , "Outgoing resets"
        .Add , , "Cumulative connections"
        '
    End With
    '
    'The system tray code
    '
    'the form must be fully visible before calling Shell_NotifyIcon
    Me.Show
    Me.Refresh
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = ImageList1.ListImages(4).Picture
        .szTip = "TCP Statistics" & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nid
    '
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '
    'this procedure receives the callbacks from the System Tray icon.
    '
    Dim Result As Long
    Dim msg As Long
    '
    'the value of X will vary depending upon the scalemode setting
    '
    If Me.ScaleMode = vbPixels Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If
    '
    Select Case msg
        Case WM_LBUTTONUP        '514 restore form window
            Me.WindowState = vbNormal
            Result = SetForegroundWindow(Me.hwnd)
            Me.Show
        Case WM_LBUTTONDBLCLK    '515 restore form window
            Me.WindowState = vbNormal
            Result = SetForegroundWindow(Me.hwnd)
            Me.Show
        Case WM_RBUTTONUP        '517 display popup menu
            Result = SetForegroundWindow(Me.hwnd)
            Me.PopupMenu Me.mPopupSys
    End Select
    
End Sub

Private Sub Form_Resize()
    '
    'this is necessary to assure that the minimized window is hidden
    '
    If Me.WindowState = vbMinimized Then Me.Hide
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '
    'this removes the icon from the system tray
    '
    Shell_NotifyIcon NIM_DELETE, nid
    '
End
End Sub

Private Sub ListView1_ItemClick(ByVal item As MSComctlLib.ListItem)
    '
    Select Case item.Index
        Case 1
            lblDesc.Caption = "Specifies the retransmission time-out (RTO) algorithm in use."
        Case 2
            lblDesc.Caption = "Specifies the minimum retransmission time-out value in milliseconds."
        Case 3
            lblDesc.Caption = "Specifies the maximum retransmission time-out value in milliseconds."
        Case 4
            lblDesc.Caption = "Specifies the maximum number of connections. If this member is -1, the maximum number of connections is dynamic."
        Case 5
            lblDesc.Caption = "Specifies the number of active opens. In an active open, the client is initiating a connection with the server."
        Case 6
            lblDesc.Caption = "Specifies the number of passive opens. In a passive open, the server is listening for a connection request from a client."
        Case 7
            lblDesc.Caption = "Specifies the number of failed connection attempts."
        Case 8
            lblDesc.Caption = "Specifies the number of established connections that have been reset."
        Case 9
            lblDesc.Caption = "Specifies the number of currently established connections."
        Case 10
            lblDesc.Caption = "Specifies the number of segments received."
        Case 11
            lblDesc.Caption = "Specifies the number of segments transmitted. This number does not include retransmitted segments."
        Case 12
            lblDesc.Caption = "Specifies the number of segments retransmitted. "
        Case 13
            lblDesc.Caption = "Specifies the number of errors received."
        Case 14
            lblDesc.Caption = "Specifies the number of segments transmitted with the reset flag set."
        Case 15
            lblDesc.Caption = "Specifies the cumulative number of connections."
    End Select
    '
End Sub

Private Sub mPopHide_Click()
    Me.Hide
    frmping.Hide
End Sub

Private Sub mPopIPResolver_Click()
frmIPResolver.Visible = True
Load frmIPResolver
End Sub

Private Sub mPopMSGSender_Click()
frmMSGSender.Visible = True
Load frmMSGSender
End Sub

Private Sub mPopOpenPorts_Click()
frmOpenPorts.Visible = True
Load frmOpenPorts
End Sub

Private Sub mPopPing_Click()
frmping.Visible = True
Load frmping
End Sub

Private Sub mPopUserNames_Click()
frmUserNames.Visible = True
Load frmUserNames
End Sub

Private Sub Timer1_Timer()
    UpdateStats
End Sub

Private Sub UpdateStats()
    '
    Dim tStats          As MIB_TCPSTATS
    Static tStaticStats As MIB_TCPSTATS
    '
    Dim lRetValue       As Long
    '
    Dim blnIsSent       As Boolean
    Dim blnIsRecv       As Boolean
    '
    lRetValue = GetTcpStatistics(tStats)
    '
    With tStats
        '
        If Not tStaticStats.dwRtoAlgorithm = .dwRtoAlgorithm Then _
            ListView1.ListItems(1).SubItems(1) = .dwRtoAlgorithm
        If Not tStaticStats.dwRtoMin = .dwRtoMin Then _
            ListView1.ListItems(2).SubItems(1) = .dwRtoMin
        If Not tStaticStats.dwRtoMax = .dwRtoMax Then _
            ListView1.ListItems(3).SubItems(1) = .dwRtoMax
        If Not tStaticStats.dwMaxConn = .dwMaxConn Then _
            ListView1.ListItems(4).SubItems(1) = .dwMaxConn
        If Not tStaticStats.dwActiveOpens = .dwActiveOpens Then _
            ListView1.ListItems(5).SubItems(1) = .dwActiveOpens
        If Not tStaticStats.dwPassiveOpens = .dwPassiveOpens Then _
            ListView1.ListItems(6).SubItems(1) = .dwPassiveOpens
        If Not tStaticStats.dwAttemptFails = .dwAttemptFails Then _
            ListView1.ListItems(7).SubItems(1) = .dwAttemptFails
        If Not tStaticStats.dwEstabResets = .dwEstabResets Then _
            ListView1.ListItems(8).SubItems(1) = .dwEstabResets
        If Not tStaticStats.dwCurrEstab = .dwCurrEstab Then _
            ListView1.ListItems(9).SubItems(1) = .dwCurrEstab
        If Not tStaticStats.dwInSegs = .dwInSegs Then _
            ListView1.ListItems(10).SubItems(1) = .dwInSegs
        If Not tStaticStats.dwOutSegs = .dwOutSegs Then _
            ListView1.ListItems(11).SubItems(1) = .dwOutSegs
        If Not tStaticStats.dwRetransSegs = .dwRetransSegs Then _
            ListView1.ListItems(12).SubItems(1) = .dwRetransSegs
        If Not tStaticStats.dwInErrs = .dwInErrs Then _
            ListView1.ListItems(13).SubItems(1) = .dwInErrs
        If Not tStaticStats.dwOutRsts = .dwOutRsts Then _
            ListView1.ListItems(14).SubItems(1) = .dwOutRsts
        If Not tStaticStats.dwNumConns = .dwNumConns Then _
            ListView1.ListItems(15).SubItems(1) = .dwNumConns
        '
    End With
    '
    blnIsRecv = (tStats.dwInSegs > tStaticStats.dwInSegs)
    blnIsSent = (tStats.dwOutSegs > tStaticStats.dwOutSegs)
    '
    If blnIsRecv And blnIsSent Then
        Set Image1.Picture = ImageList2.ListImages(4).Picture
        nid.hIcon = ImageList1.ListImages(4).Picture
    ElseIf (Not blnIsRecv) And blnIsSent Then
        Set Image1.Picture = ImageList2.ListImages(3).Picture
        nid.hIcon = ImageList1.ListImages(3).Picture
    ElseIf blnIsRecv And (Not blnIsSent) Then
        Set Image1.Picture = ImageList2.ListImages(2).Picture
        nid.hIcon = ImageList1.ListImages(2).Picture
    ElseIf Not (blnIsRecv And blnIsSent) Then
        Set Image1.Picture = ImageList2.ListImages(1).Picture
        nid.hIcon = ImageList1.ListImages(1).Picture
    End If
    '
    'Modify the system tray icon
    '
    Shell_NotifyIcon NIM_MODIFY, nid
    '
    tStaticStats = tStats
    '
End Sub


Private Sub mPopExit_Click()
    '
    'called when user clicks the popup menu Exit command
    '
    Unload Me
    End
    '
End Sub


Private Sub mPopRestore_Click()
    '
    'called when the user clicks the popup menu Restore command
    '
    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hwnd)
    Me.Show
    '
End Sub

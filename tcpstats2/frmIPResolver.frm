VERSION 5.00
Begin VB.Form frmIPResolver 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "HostName/IP Resolver"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3795
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   3795
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.ComboBox cmbHostName 
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtIPAddress 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   960
      Width           =   3615
   End
   Begin VB.CommandButton cmdGetIPAddress 
      Caption         =   "&Resolve"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Host Name/IP:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   67
      TabIndex        =   0
      Top             =   165
      Width           =   1230
   End
End
Attribute VB_Name = "frmIPResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbHostName_DropDown()
   
tmpComboStr = cmbHostName.Text
   cmbHostName.Clear
cmbHostName.Text = tmpComboStr
   Call GetServersMSGcmb(vbNullString)

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdGetIPAddress_Click()
Screen.MousePointer = vbHourglass
If Valid_IP(cmbHostName) = False Then
  txtIPAddress.Text = GetIPByHost(cmbHostName.Text)
Else
  txtIPAddress.Text = GetHostByIP(cmbHostName.Text)
End If
Screen.MousePointer = vbNormal
End Sub

Private Sub cmdGetIPAddress_GotFocus()
cmbHostName.SelStart = 0: cmbHostName.SelLength = Len(cmbHostName.Text)
End Sub

Private Sub Form_Load()
    ' Initialize the sockets library.
    InitializeSockets

    ' Display the local host's name.
    cmbHostName = LocalHostName()
End Sub
Private Sub Form_Unload(Cancel As Integer)
    ' Clean up the sockets library.
    CleanupSockets
End Sub

Private Sub cmbHostName_GotFocus()
cmbHostName.SelStart = 0: cmbHostName.SelLength = Len(cmbHostName.Text)
End Sub

Private Sub txtIPAddress_GotFocus()
txtIPAddress.SelStart = 0: txtIPAddress.SelLength = Len(txtIPAddress)
End Sub

Private Sub txtIPAddress_LostFocus()
txtIPAddress.SelLength = 0
End Sub


VERSION 5.00
Begin VB.Form frmMSGSender 
   Caption         =   "Message Sender"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   2160
      TabIndex        =   9
      Top             =   840
      Width           =   2535
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2160
      TabIndex        =   8
      Top             =   120
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2160
      TabIndex        =   7
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Height          =   1695
      Left            =   2160
      MaxLength       =   127
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1200
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Send"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Message:"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Send From:"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Send To:"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Server:"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   3000
      Width           =   4215
   End
End
Attribute VB_Name = "frmMSGSender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Combo1_DropDown()
   
tmpComboStr = Combo1.Text
   Combo1.Clear
Combo1.Text = tmpComboStr
   Call GetServersMSGcmb(vbNullString)

End Sub

Private Sub Combo2_DropDown()
   
tmpComboStr = Combo2.Text
   Combo2.Clear
Combo2.Text = tmpComboStr
   Call GetServersMSGcmb(vbNullString)

End Sub

Private Sub Combo3_DropDown()
   
tmpComboStr = Combo3.Text
   Combo3.Clear
Combo3.Text = tmpComboStr
   Call GetServersMSGcmb(vbNullString)

End Sub

Private Sub Form_Load()

   Dim tmp As String
   
  'pre-load the text boxes with
  'the local computer name for testing
   tmp = TrimNull(Space$(MAX_COMPUTERNAME + 1))
   Call GetComputerName(tmp, Len(tmp))
   
   Combo2.Text = tmp
   Combo1.Text = tmp
   Combo3.Text = tmp
   
End Sub


Private Sub Command1_Click()

   Dim msgData As NetMessageData
   Dim sSuccess As String
   
   With msgData
      .sServerName = Combo2.Text
      .sSendTo = Combo1.Text
      .sSendFrom = Combo3.Text
      .sMessage = Text4.Text
   End With
   
    sSuccess = NetSendMessage(msgData)
    
    Label1.Caption = sSuccess
    
End Sub



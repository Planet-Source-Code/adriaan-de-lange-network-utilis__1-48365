VERSION 5.00
Begin VB.Form frmUserNames 
   Caption         =   "Computer User Information"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3480
      TabIndex        =   5
      Top             =   120
      Width           =   3615
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2280
      Width           =   3615
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1560
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   840
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   3480
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmUserNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub Combo1_Click()
   
List1.Clear

   Dim tmp As String
   Dim bServername() As Byte
   
   tmp = Combo1.Text
   Call GetComputerName(tmp, Len(tmp))
   
   Combo1.Text = tmp

  'assure the server string is properly formatted
   If Len(tmp) Then
   
      If InStr(tmp, "\\") Then
            bServername = tmp & Chr$(0)
      Else: bServername = "\\" & tmp & Chr$(0)
      End If
   
   End If
   
   Combo1.Text = tmp

   Call GetUserEnumInfo(bServername())
   

End Sub

Private Sub Combo1_DropDown()
   
tmpComboStr = Combo1.Text
   Combo1.Clear
Combo1.Text = tmpComboStr
   Call GetServers(vbNullString)

End Sub

Private Sub Form_Load()

   Dim tmp As String
   Dim bServername() As Byte
   
   tmp = Space$(MAX_COMPUTERNAME + 1)
   Call GetComputerName(tmp, Len(tmp))
   
   Combo1.Text = tmp

   'tmp = "jhb_adriaand_xp" 'GetComputersName()

  'assure the server string is properly formatted
   If Len(tmp) Then
   
      If InStr(tmp, "\\") Then
            bServername = tmp & Chr$(0)
      Else: bServername = "\\" & tmp & Chr$(0)
      End If
   
   End If
   
   Combo1.Text = tmp

   Call GetUserEnumInfo(bServername())
   
End Sub



Private Sub List1_Click()

   Dim usr As USER_INFO
   Dim bUsername() As Byte
   Dim bServername() As Byte
   Dim tmp As String
  
  'This assures that both the server
  'and user params have data
   If Len(Combo1.Text) And (List1.ListIndex > -1) Then
   
      bUsername = List1.List(List1.ListIndex) & Chr$(0)
   
     'This demo uses the current machine as the
     'server param, which works on NT4 and Win2000.
     'If connected to a PDC or BDC, pass that
     'name as the server, instead of the return
     'value from GetComputerName().
      tmp = Combo1.Text
   
     'assure the server string is properly formatted
      If Len(tmp) Then
      
         If InStr(tmp, "\\") Then
               bServername = tmp & Chr$(0)
         Else: bServername = "\\" & tmp & Chr$(0)
         End If
      
      End If
   
     'Return the user information for the passed
     'user. The return values are assigned directly
     'to the non-API USER_INFO data type that we
     'defined (I prefer UDTs). Alternatively, if
     'you're a 'classy' sort of guy,  the return
     'values could be assigned directly to properties
     'in the function.
      usr = GetUserNetworkInfo(bServername(), bUsername())
      
      Text2.Text = usr.name
      
     'The call may or may not return the
     'full name, comment or usr_comment
     'members, depending on the user's
     'listing in User Manager.
      Text3.Text = usr.full_name
      Text4.Text = usr.comment
      Text5.Text = usr.usr_comment
   
   End If

End Sub




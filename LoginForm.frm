VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Hospital Mangement  - Login Screen"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8355
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LoginForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   8355
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton butCancel 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton butLogin 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Login"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   5520
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Password  :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "User Name :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   600
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   2520
      Left            =   0
      Picture         =   "LoginForm.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8400
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LNo As Byte



Private Sub butCancel_Click()
LNo = 1
End
End Sub

Private Sub butLogin_Click()

If tRS.State = 1 Then tRS.Close
tRS.Open "select * from LoginTab where uname='" & UCase(Text1) & "' and pword = '" & Text2 & "'", Conn
If tRS.EOF = True Then
MsgBox ("The entered UserName or Password is not Correct")
Text1.SetFocus
Else
MDIForm1.Show
Unload Me
'me.hide
End If

End Sub

Private Sub Command1_Click()

End Sub


Private Sub Form_Load()
Me.Height = 3015
Me.Width = 8475
StringVar = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & App.Path & "\HospitalData.mdb" & ""
Conn.ConnectionString = StringVar
Conn.Open

LNo = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
'If LNo = 1 Then End
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1)
End Sub

Private Sub Text2_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then butLogin_Click
End Sub

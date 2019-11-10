VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   10110
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   18930
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10110
   ScaleWidth      =   18930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7080
      Top             =   1200
   End
   Begin VB.Image Image1 
      Height          =   5085
      Left            =   3720
      Picture         =   "frmSplash.frx":000C
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   11835
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim I As Integer
Private Sub Timer1_Timer()
I = I + 1
If I > 4 Then
Timer1.Enabled = False
Login.Show
Unload Me

End If

End Sub

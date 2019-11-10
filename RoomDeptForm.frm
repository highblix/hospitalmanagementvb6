VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form RoomDetForm 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Room Details Form"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "RoomDeptForm.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text4 
      DataField       =   "Beds"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   2280
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      DataField       =   "Rent"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   1680
      Width           =   3135
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   240
      Top             =   3600
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Butclose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   7920
      TabIndex        =   7
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton butDelete 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton butModify 
      Caption         =   "&Update"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton butSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton ButNew 
      Caption         =   "&New"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   7920
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      DataField       =   "Facility"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   960
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      DataField       =   "RoomNo"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "No of beds"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Charges / Day"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Room No"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Details"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "RoomDetForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Butclose_Click()
Unload Me
End Sub
Private Sub butDelete_Click()
Adodc1.Recordset.Delete
End Sub

Private Sub butModify_Click()
Adodc1.Recordset.Update
End Sub

Private Sub ButNew_Click()
Adodc1.Recordset.AddNew
End Sub

Private Sub butSave_Click()
If Text1.Text = "" Then
MsgBox "Please enter all the details"
Exit Sub
End If

Adodc1.Recordset.Save
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = StringVar
Adodc1.RecordSource = "roomMainTab"
Adodc1.Refresh
End Sub

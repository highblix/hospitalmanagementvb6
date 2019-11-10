VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Hospital Management"
   ClientHeight    =   6120
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9135
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   582
      ButtonWidth     =   1138
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5745
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu MasterMenu 
      Caption         =   "Master Entry"
      Begin VB.Menu DeptDetMenu 
         Caption         =   "Dept Details"
      End
      Begin VB.Menu RoomMainMenu 
         Caption         =   "Room Main Details"
      End
   End
   Begin VB.Menu StaffMainMenu 
      Caption         =   "Staff Details"
   End
   Begin VB.Menu OutPatMenu 
      Caption         =   "Out Patient"
      Begin VB.Menu OutPatDetMenu 
         Caption         =   "Out Patient Details"
      End
      Begin VB.Menu ConChMenu 
         Caption         =   "Consultation Charges Entry"
      End
   End
   Begin VB.Menu InPatMenu 
      Caption         =   "In Patient"
      Begin VB.Menu InPatDetMenu 
         Caption         =   "In Patient Details"
      End
      Begin VB.Menu InPatadmMenu 
         Caption         =   "In Patient Admission"
      End
      Begin VB.Menu patBillMenu 
         Caption         =   "Patient Bill Entry"
      End
   End
   Begin VB.Menu RepMenu 
      Caption         =   "Report"
      Begin VB.Menu EmpListRepMenu 
         Caption         =   "Emp List Report"
      End
      Begin VB.Menu ConsRepMenu 
         Caption         =   "Consulation Report"
      End
      Begin VB.Menu BillingRepMenu 
         Caption         =   "Billing Report"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BillingRepMenu_Click()
BillRepForm.Show
End Sub

Private Sub ConChMenu_Click()
OutPatientConsulationForm.Show
End Sub

Private Sub ConsRepMenu_Click()
OutPatientListForm.Show
End Sub

Private Sub DeptDetMenu_Click()
DeptForm.Show
End Sub

Private Sub DesigMainMenu_Click()
DesignationForm.Show
End Sub

Private Sub EdMasMenu_Click()
EducationForm.Show
End Sub

Private Sub HospDetMenu_Click()
HospitalMainForm.Show
End Sub

Private Sub EmpListRepMenu_Click()
EmpListForm.Show
End Sub

Private Sub InPatadmMenu_Click()
InPatientAdmnForm.Show
End Sub

Private Sub InPatDetMenu_Click()
InPatientForm.Show
End Sub

Private Sub MDIForm_Load()
If Conn.State = 0 Then
Conn.ConnectionString = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & App.Path & "\HospitalData.mdb" & ""
Conn.Open
End If
End Sub

Private Sub MDIForm_Terminate()
End
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub

Private Sub OutPatDetMenu_Click()
OutPatientDetailsForm.Show
End Sub



Private Sub patBillMenu_Click()
BillForm.Show
End Sub

Private Sub RoomMainMenu_Click()
RoomDetForm.Show
End Sub



Private Sub StaffMainMenu_Click()
EmpDetailsForm.Show
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   3510
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   6990
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2073.824
   ScaleMode       =   0  'User
   ScaleWidth      =   6563.231
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Program will terminate in 60 seconds."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   780
      Left            =   90
      TabIndex        =   16
      Top             =   2655
      Width           =   6810
      Begin VB.Timer Timer3 
         Interval        =   1000
         Left            =   4050
         Top             =   180
      End
      Begin VB.Timer Timer2 
         Interval        =   100
         Left            =   4500
         Top             =   180
      End
      Begin MSComctlLib.ProgressBar pb 
         Height          =   285
         Left            =   90
         TabIndex        =   17
         Top             =   360
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Min             =   2
         Max             =   580
         Scrolling       =   1
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6435
      Top             =   495
   End
   Begin VB.TextBox txtTime 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4140
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1665
      Width           =   2730
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4110
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   945
      Width           =   2730
   End
   Begin VB.Frame Frame1 
      Height          =   1905
      Left            =   90
      TabIndex        =   7
      Top             =   675
      Width           =   3930
      Begin VB.TextBox txtAccLevel 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   2
         Top             =   1305
         Width           =   2325
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   795
         Width           =   2325
      End
      Begin VB.TextBox txtUserName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         TabIndex        =   0
         Top             =   270
         Width           =   2325
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   345
         Left            =   1380
         Top             =   1395
         Width           =   2325
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Access Level:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   10
         Top             =   1365
         Width           =   975
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   345
         Left            =   1380
         Top             =   885
         Width           =   2325
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   345
         Left            =   1380
         Top             =   360
         Width           =   2325
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "&Password:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   9
         Top             =   825
         Width           =   750
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "&User Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   8
         Top             =   285
         Width           =   840
      End
   End
   Begin VB.TextBox txtLogInNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4515
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   90
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
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
      Height          =   390
      Left            =   4590
      TabIndex        =   3
      Top             =   2190
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5790
      TabIndex        =   4
      Top             =   2190
      Width           =   1140
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   90
      Picture         =   "frmLogin.frx":030A
      Top             =   90
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "System Security"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   720
      TabIndex        =   15
      Top             =   135
      Width           =   1635
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Time:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   4170
      TabIndex        =   14
      Top             =   1440
      Width           =   390
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Left            =   4200
      Top             =   1755
      Width           =   2730
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   4140
      TabIndex        =   12
      Top             =   720
      Width           =   405
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Left            =   4170
      Top             =   1035
      Width           =   2730
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Left            =   4575
      Top             =   180
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Log In Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   3330
      TabIndex        =   6
      Top             =   150
      Width           =   1110
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean
Public logFile As String
Dim sqlStr As String
Dim a As Integer
Dim b
Dim x
Dim rsUserLog As DAO.Recordset
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Const MF_BYPOSITION = &H400&
Dim rsLog As DAO.Recordset

Private Sub RemoveMenus()
    Dim hMenu As Long
    hMenu = GetSystemMenu(hWnd, False)
    DeleteMenu hMenu, 6, MF_BYPOSITION
End Sub

Private Sub cmdCancel_Click()
Open logFile For Output As #1
a = a - 1
Print #1, LTrim(a)
Close #1
    LoginSucceeded = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
Set rsLog = db.OpenRecordset("SELECT * FROM Users")
rsLog.FindFirst "UserName = '" & txtUserName.Text & "'"
If Not rsLog.NoMatch Then
  If rsLog.Fields!Password = txtPassword.Text And rsLog.Fields!AccLevel = txtAccLevel.Text Then
    MsgBox "Access Granted." & Chr(13) & _
    "User Name: " & txtUserName.Text & Chr(13) & _
    "Password: " & String(Len(txtPassword.Text), "*") & Chr(13) & _
    "Date of Log In: " & txtDate.Text & Chr(13) & _
    "Time of Log In: " & txtTime.Text, vbInformation, Me.Caption
    Set rsLog = db.OpenRecordset("SELECT * FROM Users WHERE UserName = '" & txtUserName.Text & "'")
    Open userLog For Output As #1
    Print #1, rsLog.Fields!UserID
    Close #1
    Timer1.Enabled = False
    Timer2.Enabled = False
    pb.Value = 2
    x = 59
    Set rsUserLog = db.OpenRecordset("SELECT * FROM Login")
    'sqlStr = "Insert Into Login (LoginNo,UserID,LoginTime,Acclevel,Date) Values (" & txtLogInNumber.Text & ",'" & rsLog.Fields!UserID & "',#" & txtTime.Text & "#," & txtAccLevel.Text & ",#" & txtDate.Text & "#)"
    'execQuery (sqlStr)
    rsUserLog.AddNew
    rsUserLog.Fields!LoginNo = txtLogInNumber.Text
    rsUserLog.Fields!UserID = rsLog.Fields!UserID
    rsUserLog.Fields!LoginTime = txtTime.Text
    rsUserLog.Fields!AccLevel = txtAccLevel.Text
    rsUserLog.Fields!Date = txtDate.Text
    rsUserLog.Update
    frmPOSI.Show
    frmPOSI.Text1.Text = txtLogInNumber.Text
    If txtAccLevel.Text = 1 Then
    mdiPOSI.mnuCategories.Enabled = False
    mdiPOSI.mnuProducts.Enabled = False
    mdiPOSI.mnuUsers.Enabled = False
    mdiPOSI.mnuViewSales.Enabled = False
    mdiPOSI.mnuViewUsersTable.Enabled = False
    End If
    If txtAccLevel.Text = 2 Then
    mdiPOSI.mnuUsers.Enabled = False
    End If
    Unload Me
  Else
    MsgBox "Access Denied.", vbCritical, Me.Caption
  End If
Else
  MsgBox "Access Denied.", vbCritical, Me.Caption
End If
End Sub
Private Sub Form_Load()
openDB
Set rsLog = db.OpenRecordset("SELECT * FROM Users")
RemoveMenus
txtDate.Text = Format(Date, "mmmm dd, yyyy")
txtTime.Text = Format(Time, "hh:mm:ss AM/PM")
logFile = App.Path & "\LogFile"
Open logFile For Input As #1
If Not EOF(1) Then
Line Input #1, b
a = Val(b)
End If
Close #1
Open logFile For Output As #1
a = a + 1
Print #1, LTrim(a)
Close #1
txtLogInNumber.Text = Format(a, "#0000#")
x = 59
userLog = App.Path & "\userLogFile"
End Sub
Private Sub Timer1_Timer()
txtDate.Text = Format(Date, "mmmm dd, yyyy")
txtTime.Text = Format(Time, "hh:mm:ss AM/PM")
End Sub
Private Sub Timer2_Timer()
pb.Value = pb.Value + 1
If pb.Value = 580 Then
MsgBox "Your time is up!" & Chr(13) & "Click <OK> to end the program.", vbCritical, Me.Caption
Open logFile For Output As #1
a = a - 1
Print #1, LTrim(a)
Close #1
End
End If
End Sub
Private Sub Timer3_Timer()
Frame2.Caption = "Program will terminate in " & x & " seconds."
x = x - 1
End Sub

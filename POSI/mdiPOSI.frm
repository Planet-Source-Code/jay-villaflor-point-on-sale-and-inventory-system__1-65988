VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiPOSI 
   BackColor       =   &H8000000C&
   Caption         =   "Point On-Sale & Inventory System"
   ClientHeight    =   5445
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7650
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   135
      Top             =   225
   End
   Begin MSComctlLib.StatusBar st1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   5205
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   7056
            MinWidth        =   7056
            Text            =   "Point On-Sale & Inventory System"
            TextSave        =   "Point On-Sale & Inventory System"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
            Text            =   "Total Sales: >>"
            TextSave        =   "Total Sales: >>"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   4057
            MinWidth        =   4057
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuDatabase 
      Caption         =   "&Database"
      Begin VB.Menu mnuProducts 
         Caption         =   "&Products"
      End
      Begin VB.Menu mnuCategories 
         Caption         =   "&Categories"
      End
      Begin VB.Menu mnuUsers 
         Caption         =   "&Users"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewUsersTable 
         Caption         =   "&User's Table"
      End
      Begin VB.Menu mnuViewSales 
         Caption         =   "&Sales' Table"
      End
   End
End
Attribute VB_Name = "mdiPOSI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Const MF_BYPOSITION = &H400&

Private Sub MDIForm_Load()
RemoveMenus
st1.Panels(2).Text = "Time: " & Format(Time, "hh:mm:ss AM/PM")
st1.Panels(3).Text = "Date: " & Date
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Cancel = 1
End Sub

Private Sub mnuCategories_Click()
frmCategories.ZOrder (0)
frmCategories.Show
End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnuProducts_Click()
frmProducts.ZOrder (0)
frmProducts.Show
End Sub

Private Sub mnuUsers_Click()
frmUsers.ZOrder (0)
frmUsers.Show
End Sub

Private Sub mnuViewSales_Click()
frmSales.Show
End Sub

Private Sub mnuViewUsersTable_Click()
frmUserLog.Show
End Sub

Private Sub RemoveMenus()
    Dim hMenu As Long
    hMenu = GetSystemMenu(hWnd, False)
    DeleteMenu hMenu, 6, MF_BYPOSITION
End Sub

Private Sub Timer1_Timer()
st1.Panels(2).Text = "Time: " & Format(Time, "hh:mm:ss AM/PM")
st1.Panels(3).Text = "Date: " & Date
End Sub

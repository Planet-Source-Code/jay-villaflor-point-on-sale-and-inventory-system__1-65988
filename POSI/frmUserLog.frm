VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmUserLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Log in Table"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   8925
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   7470
      TabIndex        =   1
      Top             =   5670
      Width           =   1365
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   5505
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   9710
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Log In No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "User ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Log in Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Log Out Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Access Level"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Log In Date"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmUserLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsUserLog As DAO.Recordset

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Top = 200
Me.Left = 200
openDB
Set rsUserLog = db.OpenRecordset("SELECT * FROM Login")
Call loadLV
End Sub

Private Sub loadLV()
On Error Resume Next
lv1.ListItems.Clear
Do While Not rsUserLog.EOF
Set x = lv1.ListItems.Add(, , rsUserLog.Fields!LoginNo)
x.SubItems(1) = rsUserLog.Fields!UserID
x.SubItems(2) = rsUserLog.Fields!LoginTime
x.SubItems(3) = rsUserLog.Fields!LogOut
x.SubItems(4) = rsUserLog.Fields!AccLevel
x.SubItems(5) = rsUserLog.Fields!Date
rsUserLog.MoveNext
Loop
End Sub

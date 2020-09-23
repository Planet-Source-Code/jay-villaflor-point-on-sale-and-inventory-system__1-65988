VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Total Sales"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
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
   ScaleHeight     =   5955
   ScaleWidth      =   10380
   Begin MSComCtl2.MonthView mv1 
      Height          =   2370
      Left            =   7605
      TabIndex        =   7
      Top             =   405
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   24444929
      CurrentDate     =   37670
   End
   Begin VB.TextBox txttotalSales 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7470
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3240
      Width           =   2805
   End
   Begin VB.TextBox txtSales 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7470
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4365
      Width           =   2805
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   420
      Left            =   8955
      TabIndex        =   1
      Top             =   5490
      Width           =   1365
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   5865
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   10345
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Sales ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Total Sales"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Time Generated"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date Generated"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "User ID"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblTotalSales 
      AutoSize        =   -1  'True
      Caption         =   "Total Sales for"
      Height          =   195
      Left            =   7470
      TabIndex        =   6
      Top             =   2925
      Width           =   1035
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   285
      Left            =   7515
      Top             =   3285
      Width           =   2805
   End
   Begin VB.Label Label2 
      Caption         =   "Select Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7650
      TabIndex        =   4
      Top             =   90
      Width           =   2355
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Overall Sales:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7470
      TabIndex        =   3
      Top             =   4050
      Width           =   1200
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   7515
      Top             =   4410
      Width           =   2805
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSales As DAO.Recordset
Dim rs1 As DAO.Recordset
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Top = 200
Me.Left = 200
openDB
Set rsSales = db.OpenRecordset("SELECT * FROM Sales")
Call loadLV
Set rs1 = db.OpenRecordset("SELECT SUM(TotalSales) as totSales FROM Sales")
txtSales.Text = Format(rs1.Fields!totSales, "###,###,###.#0")
mv1.Value = Date
End Sub
Private Sub loadLV()
lv1.ListItems.Clear
Do While Not rsSales.EOF
Set x = lv1.ListItems.Add(, , rsSales.Fields(0))
x.SubItems(1) = rsSales.Fields(1)
x.SubItems(2) = rsSales.Fields(2)
x.SubItems(3) = rsSales.Fields(3)
x.SubItems(4) = rsSales.Fields(4)
rsSales.MoveNext
Loop
End Sub
Private Sub mv1_DateClick(ByVal DateClicked As Date)
Set rsSales = db.OpenRecordset("SELECT * FROM Sales WHERE DateGenerated = #" & DateClicked & "#")
Call loadLV
If lv1.ListItems.Count <> 0 Then
lblTotalSales.Caption = "Total Sales for " & DateClicked
txttotalSales.Text = Format(lv1.ListItems.Item(1).SubItems(1), "###,###,###.#0")
Else
lblTotalSales.Caption = "Total Sales for " & DateClicked
txttotalSales.Text = ""
End If
End Sub

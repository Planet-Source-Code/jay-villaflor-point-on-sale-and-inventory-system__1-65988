VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmViewer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Products Viewer"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11820
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
   ScaleHeight     =   7200
   ScaleWidth      =   11820
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   11715
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   330
         Left            =   8865
         TabIndex        =   7
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   330
         Left            =   10260
         TabIndex        =   6
         Top             =   225
         Width           =   1365
      End
      Begin VB.ComboBox cboTemp 
         Height          =   315
         ItemData        =   "frmViewer.frx":0000
         Left            =   5085
         List            =   "frmViewer.frx":0002
         TabIndex        =   5
         Top             =   225
         Width           =   2130
      End
      Begin VB.ComboBox cboView 
         Height          =   315
         ItemData        =   "frmViewer.frx":0004
         Left            =   1215
         List            =   "frmViewer.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   225
         Width           =   2130
      End
      Begin VB.Label lblTemp 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   4050
         TabIndex        =   4
         Top             =   270
         Width           =   585
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   5130
         Top             =   315
         Width           =   2130
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vew By:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   270
         Width           =   645
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   1260
         Top             =   315
         Width           =   2130
      End
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   6360
      Left            =   45
      TabIndex        =   0
      Top             =   765
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   11218
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   5116
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Category ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Category Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Quantity"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Unit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Unit Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Cost Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "User ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Entry Date"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsProducts As DAO.Recordset
Dim rsCategories As DAO.Recordset

Private Sub cboTemp_Change()
If cboView.Text = "Description" Then
Set rsProducts = db.OpenRecordset("SELECT Products.ProdCode, Products.ProdDescription, Products.Qty, Products.Unit, Products.UnitPrice, Products.CostPrice, Products.UserID, Products.EntryDate, Products.CategoryID, Categories.CategoryName FROM Categories INNER JOIN Products ON Categories.CategoryID = Products.CategoryID WHERE Products.ProdDescription LIKE '" & cboTemp.Text & "*'")
lv1.ListItems.Clear
Call loadLV
Else
Set rsProducts = db.OpenRecordset("SELECT Products.ProdCode, Products.ProdDescription, Products.Qty, Products.Unit, Products.UnitPrice, Products.CostPrice, Products.UserID, Products.EntryDate, Products.CategoryID, Categories.CategoryName FROM Categories INNER JOIN Products ON Categories.CategoryID = Products.CategoryID WHERE Categories.CategoryName LIKE '" & cboTemp.Text & "*'")
lv1.ListItems.Clear
Call loadLV
End If
End Sub

Private Sub cboTemp_Click()
If cboView.Text = "Description" Then
Set rsProducts = db.OpenRecordset("SELECT Products.ProdCode, Products.ProdDescription, Products.Qty, Products.Unit, Products.UnitPrice, Products.CostPrice, Products.UserID, Products.EntryDate, Products.CategoryID, Categories.CategoryName FROM Categories INNER JOIN Products ON Categories.CategoryID = Products.CategoryID WHERE Products.ProdDescription LIKE '" & cboTemp.Text & "*'")
lv1.ListItems.Clear
Call loadLV
Else
Set rsProducts = db.OpenRecordset("SELECT Products.ProdCode, Products.ProdDescription, Products.Qty, Products.Unit, Products.UnitPrice, Products.CostPrice, Products.UserID, Products.EntryDate, Products.CategoryID, Categories.CategoryName FROM Categories INNER JOIN Products ON Categories.CategoryID = Products.CategoryID WHERE Categories.CategoryName LIKE '" & cboTemp.Text & "*'")
lv1.ListItems.Clear
Call loadLV
End If
End Sub

Private Sub cboView_Click()
lblTemp.Caption = cboView.Text & ":"
If cboView.Text = "Category" Then
Set rsCategories = db.OpenRecordset("SELECT * FROM Categories")
cboTemp.Clear
Do While Not rsCategories.EOF
cboTemp.AddItem rsCategories.Fields!CategoryName
rsCategories.MoveNext
Loop
Else
Set rsProducts = db.OpenRecordset("SELECT * FROM Products")
cboTemp.Clear
Do While Not rsProducts.EOF
cboTemp.AddItem rsProducts.Fields!ProdDescription
rsProducts.MoveNext
Loop
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdRefresh_Click()
cboTemp.Text = ""
cboView.ListIndex = -1
End Sub

Private Sub Form_Load()
Me.Top = 25
Me.Left = 25
openDB
Set rsProducts = db.OpenRecordset("SELECT Products.ProdCode, Products.ProdDescription, Products.Qty, Products.Unit, Products.UnitPrice, Products.CostPrice, Products.UserID, Products.EntryDate, Products.CategoryID, Categories.CategoryName FROM Categories INNER JOIN Products ON Categories.CategoryID = Products.CategoryID")
Call loadLV
End Sub

Private Sub loadLV()
Do While Not rsProducts.EOF
Set X = lv1.ListItems.Add(, , rsProducts.Fields!ProdCode)
X.SubItems(1) = rsProducts.Fields!ProdDescription
X.SubItems(2) = rsProducts.Fields!CategoryID
X.SubItems(3) = rsProducts.Fields!CategoryName
X.SubItems(4) = rsProducts.Fields!Qty
X.SubItems(5) = rsProducts.Fields!Unit
X.SubItems(6) = rsProducts.Fields!UnitPrice
X.SubItems(7) = rsProducts.Fields!CostPrice
X.SubItems(8) = rsProducts.Fields!UserID
X.SubItems(9) = rsProducts.Fields!EntryDate
rsProducts.MoveNext
Loop
End Sub

Private Sub lv1_DblClick()
frmPOSI.txtProdCode.Text = lv1.ListItems.Item(lv1.SelectedItem.Index).Text
Unload Me
End Sub

Private Sub lv1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 13
        frmPOSI.txtProdCode.Text = lv1.ListItems.Item(lv1.SelectedItem.Index).Text
        Unload Me
End Select
End Sub

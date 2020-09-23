VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmProducts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Product Entry"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   330
      Left            =   4770
      TabIndex        =   12
      Top             =   3510
      Width           =   1140
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2475
      MaxLength       =   50
      TabIndex        =   13
      Top             =   3510
      Width           =   2085
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   3030
      Left            =   90
      TabIndex        =   28
      Top             =   3915
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   5345
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   5115
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Quantity"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Unit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Unit Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cost Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Entry Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Category ID"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   420
      Left            =   5805
      TabIndex        =   11
      Top             =   2880
      Width           =   1230
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   420
      Left            =   4545
      TabIndex        =   8
      Top             =   2430
      Width           =   1230
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   420
      Left            =   5805
      TabIndex        =   10
      Top             =   2430
      Width           =   1230
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   420
      Left            =   4545
      TabIndex        =   9
      Top             =   2880
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3390
      Left            =   90
      TabIndex        =   14
      Top             =   45
      Width           =   7080
      Begin VB.TextBox txtCategoryName 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1485
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2925
         Width           =   2085
      End
      Begin VB.ComboBox cboCategory 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1350
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2475
         Width           =   1410
      End
      Begin VB.TextBox txtEntryDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2115
         Width           =   2760
      End
      Begin VB.TextBox txtCostPrice 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4050
         TabIndex        =   5
         Top             =   1755
         Width           =   1410
      End
      Begin VB.TextBox txtUnitPrice 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1350
         TabIndex        =   4
         Top             =   1755
         Width           =   1410
      End
      Begin VB.TextBox txtUnit 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4050
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1395
         Width           =   2670
      End
      Begin VB.TextBox txtUserID 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   4590
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   225
         Width           =   2130
      End
      Begin VB.TextBox txtProdCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1350
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   675
         Width           =   2130
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1035
         Width           =   5370
      End
      Begin VB.TextBox txtQuantity 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1350
         TabIndex        =   2
         Top             =   1395
         Width           =   1410
      End
      Begin VB.Shape Shape10 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   1530
         Top             =   2970
         Width           =   2085
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Category Name:"
         Height          =   195
         Left            =   180
         TabIndex        =   27
         Top             =   2970
         Width           =   1185
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Category ID:"
         Height          =   195
         Left            =   180
         TabIndex        =   25
         Top             =   2520
         Width           =   945
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   1395
         Top             =   2565
         Width           =   1410
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Entry Date:"
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Top             =   2160
         Width           =   840
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   1395
         Top             =   2160
         Width           =   2760
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cost Price:"
         Height          =   195
         Left            =   3150
         TabIndex        =   23
         Top             =   1800
         Width           =   780
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   4095
         Top             =   1800
         Width           =   1410
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Unit Price:"
         Height          =   195
         Left            =   180
         TabIndex        =   22
         Top             =   1800
         Width           =   735
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   1395
         Top             =   1800
         Width           =   1410
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Unit:"
         Height          =   195
         Left            =   3580
         TabIndex        =   21
         Top             =   1440
         Width           =   345
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   4095
         Top             =   1440
         Width           =   2670
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Product Entry"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   180
         TabIndex        =   20
         Top             =   180
         Width           =   2535
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "User ID:"
         Height          =   195
         Left            =   3690
         TabIndex        =   19
         Top             =   270
         Width           =   600
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   4635
         Top             =   270
         Width           =   2130
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Product Code:"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   720
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   1395
         Top             =   720
         Width           =   2130
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Description:"
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   1080
         Width           =   855
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   1395
         Top             =   1080
         Width           =   5370
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Quantity:"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   1440
         Width           =   690
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   1395
         Top             =   1440
         Width           =   1410
      End
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Search for Product Description:"
      Height          =   195
      Left            =   135
      TabIndex        =   29
      Top             =   3555
      Width           =   2250
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   285
      Left            =   2520
      Top             =   3555
      Width           =   2085
   End
End
Attribute VB_Name = "frmProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCategory As DAO.Recordset
Dim rsProducts As DAO.Recordset
Dim txt As Control
Dim temp
Dim esc As Byte
Dim productFile As String
Dim a As Integer
Dim b

Private Sub cboCategory_Change()
On Error Resume Next
Set rsCategory = db.OpenRecordset("SELECT * FROM Categories WHERE CategoryID = '" & cboCategory.Text & "'")
txtCategoryName.Text = rsCategory.Fields!CategoryName
End Sub

Private Sub cboCategory_Click()
On Error Resume Next
Set rsCategory = db.OpenRecordset("SELECT * FROM Categories WHERE CategoryID = '" & cboCategory.Text & "'")
txtCategoryName.Text = rsCategory.Fields!CategoryName
End Sub


Private Sub cmdAdd_Click()
On Error GoTo errHandler
If cmdAdd.Caption = "&Add" Then
rsProducts.AddNew
Frame1.Enabled = True
cmdDelete.Enabled = False
cmdUpdate.Enabled = False
Call clrTxt
txtDescription.SetFocus
cmdAdd.Caption = "&Save"
esc = 1
Open productFile For Output As #1
a = a + 1
Print #1, LTrim(a)
Close #1
txtProdCode.Text = a
txtEntryDate.Text = Format(Date, "mmmm dd, yyyy")
txtUserID.Text = temp
Else
 If txtProdCode.Text <> "" And txtDescription.Text <> "" _
 And txtQuantity.Text <> "" And txtUnit.Text <> "" And txtUnitPrice.Text <> "" _
 And txtCostPrice.Text <> "" And txtEntryDate.Text <> "" _
 And cboCategory.Text <> "" Then
 If IsNumeric(txtUnitPrice.Text) = False Then
 MsgBox "Unit price must be of numeric type.", vbInformation, Me.Caption
 txtUnitPrice.SetFocus
 Exit Sub
 End If
 If IsNumeric(txtCostPrice.Text) = False Then
 MsgBox "Cost price must be of numeric type.", vbInformation, Me.Caption
 txtCostPrice.SetFocus
 Exit Sub
 End If
 If Val(txtUnitPrice.Text) >= Val(txtCostPrice.Text) Then
 MsgBox "Cost price must be greater than Unit price.", vbInformation, Me.Caption
 txtCostPrice.SetFocus
 Exit Sub
 End If
 Call ConFields
 rsProducts.Update
 rsProducts.MoveFirst
 Call loadLV
 MsgBox "New product has been added.", vbInformation, Me.Caption
 Call clrTxt
 cmdAdd.Caption = "&Add"
 Frame1.Enabled = False
 cmdDelete.Enabled = True
 cmdUpdate.Enabled = True
 esc = 0
 Else
 MsgBox "Fill up all the required fields.", vbInformation, Me.Caption
 Exit Sub
 End If
End If
Exit Sub
errHandler:
MsgBox Err.Description, vbInformation, Me.Caption
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
If txtProdCode.Text <> "" Then
Set rsProducts = db.OpenRecordset("SELECT * FROM Products WHERE ProdCode = '" & txtProdCode.Text & "'")
reply = MsgBox("Product Code: " & rsProducts.Fields!ProdCode & Chr(13) & "Description: " & rsProducts.Fields!ProdDescription & Chr(13) & "Are you sure you want to Delete this product?", vbQuestion + vbYesNo, Me.Caption)
If reply = vbYes Then
rsProducts.Delete
Call clrTxt
Set rsProducts = db.OpenRecordset("SELECT * FROM Products")
Call loadLV
cmdAdd.Enabled = True
Frame1.Enabled = False
MsgBox "Product deleted.", vbInformation, Me.Caption
End If
Else
MsgBox "No product to delete.", vbInformation, Me.Caption
End If
End Sub

Private Sub cmdRefresh_Click()
txtSearch.Text = ""
End Sub

Private Sub cmdUpdate_Click()
 If txtProdCode.Text <> "" And txtDescription.Text <> "" _
 And txtQuantity.Text <> "" And txtUnit.Text <> "" And txtUnitPrice.Text <> "" _
 And txtCostPrice.Text <> "" And txtEntryDate.Text <> "" _
 And cboCategory.Text <> "" Then
 If IsNumeric(txtUnitPrice.Text) = False Then
 MsgBox "Unit price must be of numeric type.", vbInformation, Me.Caption
 txtUnitPrice.SetFocus
 Exit Sub
 End If
 If IsNumeric(txtCostPrice.Text) = False Then
 MsgBox "Cost price must be of numeric type.", vbInformation, Me.Caption
 txtCostPrice.SetFocus
 Exit Sub
 End If
 If Val(txtUnitPrice.Text) >= Val(txtCostPrice.Text) Then
 MsgBox "Cost price must be greater than Unit price.", vbInformation, Me.Caption
 txtCostPrice.SetFocus
 Exit Sub
 End If
Call updateRec
Set rsProducts = db.OpenRecordset("SELECT * FROM Products")
Frame1.Enabled = False
cmdAdd.Enabled = True
Call clrTxt
Call loadLV
MsgBox "Product updated", vbInformation, Me.Caption
 Else
 MsgBox "Fill up all the required fields.", vbInformation, Me.Caption
 Exit Sub
 End If
End Sub

Private Sub Form_Activate()
Set rsCategory = db.OpenRecordset("SELECT * FROM Categories")
Call cboAddItem
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 13
        SendKeys "{tab}"
    Case vbKeyEscape
        If esc = 1 Then
            reply = MsgBox("Cancel operation?", vbQuestion + vbYesNo, Me.Caption)
            If reply = vbYes Then
            rsProducts.CancelUpdate
            Call clrTxt
            cmdAdd.Caption = "&Add"
            Frame1.Enabled = False
            cmdDelete.Enabled = True
            cmdUpdate.Enabled = True
            Open productFile For Output As #1
            a = a - 1
            Print #1, LTrim(a)
            Close #1
            esc = 0
            End If
        End If
End Select
End Sub

Private Sub Form_Load()
Me.Top = 100
Me.Left = 100
Me.WindowState = 0
esc = 0
Call clrTxt
openDB
Set rsCategory = db.OpenRecordset("SELECT * FROM Categories")
Call cboAddItem
userLog = App.Path & "\userLogFile"
Open userLog For Input As #1
If Not EOF(1) Then
Line Input #1, temp
End If
Close #1
txtUserID.Text = temp
productFile = App.Path & "\Products"
Open productFile For Input As #1
If Not EOF(1) Then
Line Input #1, b
a = Val(b)
End If
Close #1
Set rsProducts = db.OpenRecordset("SELECT * FROM Products")
Call loadLV
End Sub

Private Sub cboAddItem()
cboCategory.Clear
Do While Not rsCategory.EOF
cboCategory.AddItem rsCategory.Fields!CategoryID
rsCategory.MoveNext
Loop
End Sub

Private Sub clrTxt()
For Each txt In Me.Controls
If TypeOf txt Is TextBox Then
txt.Text = ""
End If
Next
cboCategory.ListIndex = -1
End Sub

Private Sub loadLV()
lv1.ListItems.Clear
Do While Not rsProducts.EOF
Set k = lv1.ListItems.Add(, , rsProducts.Fields!ProdCode)
k.SubItems(1) = rsProducts.Fields!ProdDescription
k.SubItems(2) = rsProducts.Fields!Qty
k.SubItems(3) = rsProducts.Fields!Unit
k.SubItems(4) = rsProducts.Fields!UnitPrice
k.SubItems(5) = rsProducts.Fields!CostPrice
k.SubItems(6) = rsProducts.Fields!EntryDate
k.SubItems(7) = rsProducts.Fields!CategoryID
rsProducts.MoveNext
Loop
End Sub

Private Sub ConFields()
rsProducts.Fields!UserID = txtUserID.Text
rsProducts.Fields!ProdCode = txtProdCode.Text
rsProducts.Fields!ProdDescription = txtDescription.Text
rsProducts.Fields!Qty = txtQuantity.Text
rsProducts.Fields!Unit = txtUnit.Text
rsProducts.Fields!UnitPrice = txtUnitPrice.Text
rsProducts.Fields!CostPrice = txtCostPrice.Text
rsProducts.Fields!EntryDate = txtEntryDate.Text
rsProducts.Fields!CategoryID = cboCategory.Text
End Sub

Private Sub RetFields()
txtUserID.Text = rsProducts.Fields!UserID
txtProdCode.Text = rsProducts.Fields!ProdCode
txtDescription.Text = rsProducts.Fields!ProdDescription
txtQuantity.Text = rsProducts.Fields!Qty
txtUnit.Text = rsProducts.Fields!Unit
txtUnitPrice.Text = rsProducts.Fields!UnitPrice
txtCostPrice.Text = rsProducts.Fields!CostPrice
txtEntryDate.Text = rsProducts.Fields!EntryDate
cboCategory.Text = rsProducts.Fields!CategoryID
End Sub

Private Sub lv1_DblClick()
If lv1.ListItems.Count <> 0 Then
Set rsProducts = db.OpenRecordset("SELECT * FROM Products WHERE ProdCode = '" & lv1.ListItems.Item(lv1.SelectedItem.Index).Text & "'")
Call RetFields
Frame1.Enabled = True
txtDescription.SetFocus
SendKeys "{home}+{end}"
cmdAdd.Enabled = False
End If
End Sub

Private Sub updateRec()
sqlStr = "UPDATE Products SET " _
& " ProdCode = '" & txtProdCode.Text & "'" _
& ", ProdDescription = '" & txtDescription.Text & "'" _
& ", Qty = '" & txtQuantity.Text & "'" _
& ", Unit = '" & txtUnit.Text & "'" _
& ", UnitPrice = '" & txtUnitPrice.Text & "'" _
& ", CostPrice = '" & txtCostPrice.Text & "'" _
& ", UserID = '" & txtUserID.Text & "'" _
& ", EntryDate = '" & txtEntryDate.Text & "'" _
& ", CategoryID = '" & cboCategory.Text & "' WHERE ProdCode = '" & txtProdCode.Text & "'"
execQuery (sqlStr)
End Sub


Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
If Not KeyAscii = vbKeyBack Then
KeyAscii = 0
Else
KeyAscii = vbKeyBack
End If
End If
End Sub

Private Sub txtSearch_Change()
Set rsProducts = db.OpenRecordset("SELECT * FROM Products WHERE ProdDescription Like '" & txtSearch.Text & "*'")
Call loadLV
End Sub

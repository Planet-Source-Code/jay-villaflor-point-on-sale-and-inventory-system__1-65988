VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCategories 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Categories"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   7080
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5895
      TabIndex        =   13
      Top             =   2295
      Width           =   1095
   End
   Begin VB.TextBox txtSearch 
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
      Height          =   285
      Left            =   3645
      TabIndex        =   11
      Top             =   2295
      Width           =   2130
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   3480
      Left            =   90
      TabIndex        =   10
      ToolTipText     =   "Double click an item to edit or delete."
      Top             =   2700
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   6138
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Category ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Category Name"
         Object.Width           =   9526
      EndProperty
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3105
      TabIndex        =   4
      Top             =   1710
      Width           =   1230
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4365
      TabIndex        =   5
      Top             =   1710
      Width           =   1230
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1845
      TabIndex        =   3
      Top             =   1710
      Width           =   1230
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5625
      TabIndex        =   6
      Top             =   1710
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
      Height          =   2085
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   6900
      Begin VB.TextBox txtCategoryName 
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
         Height          =   285
         Left            =   1350
         TabIndex        =   2
         Top             =   1035
         Width           =   5370
      End
      Begin VB.TextBox txtCategoryID 
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
         Height          =   285
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   675
         Width           =   2130
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   1395
         Top             =   1080
         Width           =   5370
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
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
         Left            =   180
         TabIndex        =   9
         Top             =   1080
         Width           =   465
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   1395
         Top             =   720
         Width           =   2130
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Category ID:"
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
         Left            =   180
         TabIndex        =   8
         Top             =   720
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Categories"
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
         TabIndex        =   7
         Top             =   180
         Width           =   1950
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Search for Category Name:"
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
      Left            =   1530
      TabIndex        =   12
      Top             =   2340
      Width           =   1980
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   285
      Left            =   3690
      Top             =   2340
      Width           =   2130
   End
End
Attribute VB_Name = "frmCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCategory As DAO.Recordset
Dim txt As Control
Dim esc As Byte
Dim categoryPath As String
Dim a As Integer
Dim b
Dim tempCategory As String

Private Sub cmdAdd_Click()
On Error GoTo errHandler
If cmdAdd.Caption = "&Add" Then
rsCategory.AddNew
Frame1.Enabled = True
cmdDelete.Enabled = False
cmdUpdate.Enabled = False
Call clrTxt
txtCategoryName.SetFocus
cmdAdd.Caption = "&Save"
esc = 1
Open categoryPath For Output As #1
a = a + 1
Print #1, LTrim(a)
Close #1
txtCategoryID.Text = Format(a, "000#")
Else
 If txtCategoryID.Text <> "" And txtCategoryName.Text <> "" Then
 Call ConFields
 rsCategory.Update
 rsCategory.MoveFirst
 Call loadLV
 MsgBox "New category has been added.", vbInformation, Me.Caption
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
If txtCategoryID.Text <> "" Then
Set rsCategory = db.OpenRecordset("SELECT * FROM Categories WHERE CategoryID = '" & txtCategoryID.Text & "'")
reply = MsgBox("Category ID: " & rsCategory.Fields!CategoryID & Chr(13) & "Category Name: " & rsCategory.Fields!CategoryName & Chr(13) & "Are you sure you want to Delete this category?", vbQuestion + vbYesNo, Me.Caption)
If reply = vbYes Then
rsCategory.Delete
Call clrTxt
Set rsCategory = db.OpenRecordset("SELECT * FROM Categories")
Call loadLV
cmdAdd.Enabled = True
Frame1.Enabled = False
MsgBox "Category deleted.", vbInformation, Me.Caption
End If
Else
MsgBox "No category to delete.", vbInformation, Me.Caption
End If
End Sub

Private Sub cmdRefresh_Click()
txtSearch.Text = ""
End Sub

Private Sub cmdUpdate_Click()
If txtCategoryID.Text <> "" And txtCategoryName.Text <> "" Then
Call updateRec
Set rsCategory = db.OpenRecordset("SELECT * FROM Categories")
Frame1.Enabled = False
cmdAdd.Enabled = True
Call clrTxt
Call loadLV
MsgBox "Category updated", vbInformation, Me.Caption
 Else
 MsgBox "Fill up all the required fields.", vbInformation, Me.Caption
 Exit Sub
 End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 13
        SendKeys "{tab}"
    Case vbKeyEscape
        If esc = 1 Then
            reply = MsgBox("Cancel operation?", vbQuestion + vbYesNo, Me.Caption)
            If reply = vbYes Then
            rsCategory.CancelUpdate
            Call clrTxt
            cmdAdd.Caption = "&Add"
            Frame1.Enabled = False
            cmdDelete.Enabled = True
            cmdUpdate.Enabled = True
            Open categoryPath For Output As #1
            a = a - 1
            Print #1, LTrim(a)
            Close #1
            esc = 0
            End If
        End If
End Select
End Sub

Private Sub Form_Load()
Me.Top = 200
Me.Left = 200
Me.WindowState = 0
openDB
Set rsCategory = db.OpenRecordset("SELECT * FROM Categories")
esc = 0
categoryPath = App.Path & "\Categories"
Open categoryPath For Input As #1
If Not EOF(1) Then
Line Input #1, b
a = Val(b)
End If
Close #1
Call loadLV
End Sub

Private Sub clrTxt()
For Each txt In Me.Controls
If TypeOf txt Is TextBox Then
txt.Text = ""
End If
Next
End Sub

Private Sub ConFields()
rsCategory.Fields!CategoryID = txtCategoryID.Text
rsCategory.Fields!CategoryName = txtCaps(txtCategoryName.Text)
End Sub

Private Sub RetFields()
txtCategoryID.Text = rsCategory.Fields!CategoryID
txtCategoryName.Text = rsCategory.Fields!CategoryName
End Sub

Private Sub updateRec()
sqlStr = "UPDATE Categories SET " _
& " CategoryID = '" & txtCategoryID.Text & "'" _
& ", CategoryName = '" & txtCaps(txtCategoryName.Text) & "' WHERE CategoryID = '" & txtCategoryID.Text & "'"
execQuery (sqlStr)
End Sub

Private Sub loadLV()
lv1.ListItems.Clear
Do While Not rsCategory.EOF
Set j = lv1.ListItems.Add(, , rsCategory.Fields!CategoryID)
j.SubItems(1) = rsCategory.Fields!CategoryName
rsCategory.MoveNext
Loop
End Sub

Private Sub lv1_DblClick()
If lv1.ListItems.Count <> 0 Then
Set rsCategory = db.OpenRecordset("SELECT * FROM Categories WHERE CategoryID = '" & lv1.ListItems.Item(lv1.SelectedItem.Index).Text & "'")
Call RetFields
Frame1.Enabled = True
txtCategoryName.SetFocus
SendKeys "{home}+{end}"
cmdAdd.Enabled = False
End If
End Sub

Private Sub txtSearch_Change()
Set rsCategory = db.OpenRecordset("SELECT * FROM Categories WHERE CategoryName Like '" & txtSearch.Text & "*'")
Call loadLV
End Sub

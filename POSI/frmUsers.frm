VERSION 5.00
Begin VB.Form frmUsers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User's Information"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   7080
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   420
      Left            =   5625
      TabIndex        =   8
      Top             =   2115
      Width           =   1230
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   420
      Left            =   4365
      TabIndex        =   5
      Top             =   1665
      Width           =   1230
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Search"
      Height          =   420
      Left            =   5625
      TabIndex        =   7
      Top             =   1665
      Width           =   1230
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   420
      Left            =   4365
      TabIndex        =   6
      Top             =   2115
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
      Height          =   2625
      Left            =   90
      TabIndex        =   9
      Top             =   45
      Width           =   6900
      Begin VB.ComboBox cboAccLevel 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmUsers.frx":0000
         Left            =   1350
         List            =   "frmUsers.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1755
         Width           =   1410
      End
      Begin VB.TextBox txtPosition 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1350
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1395
         Width           =   1410
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1035
         Width           =   5370
      End
      Begin VB.TextBox txtUserName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1350
         MaxLength       =   15
         TabIndex        =   0
         Top             =   675
         Width           =   2130
      End
      Begin VB.TextBox txtUserID 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   4590
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   225
         Width           =   2130
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1350
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   2160
         Width           =   1410
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   330
         Left            =   1395
         Top             =   1800
         Width           =   1410
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   1395
         Top             =   1440
         Width           =   1410
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Position:"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   1440
         Width           =   615
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
         Height          =   195
         Left            =   180
         TabIndex        =   16
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
         Caption         =   "User Name:"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   720
         Width           =   840
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   4635
         Top             =   270
         Width           =   2130
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "User ID:"
         Height          =   195
         Left            =   3690
         TabIndex        =   14
         Top             =   270
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "User's Information"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   180
         TabIndex        =   13
         Top             =   180
         Width           =   2970
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Access Level:"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   1800
         Width           =   975
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   1395
         Top             =   2205
         Width           =   1410
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   2205
         Width           =   750
      End
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsUser As DAO.Recordset
Dim txt As Control
Dim esc As Byte
Dim userPath As String
Dim a As Integer
Dim b
Dim tempUser As String

Private Sub cmdAdd_Click()
On Error GoTo errHandler
If cmdAdd.Caption = "&Add" Then
rsUser.AddNew
Frame1.Enabled = True
cmdDelete.Enabled = False
cmdUpdate.Enabled = False
Call clrTxt
txtUserName.SetFocus
cmdAdd.Caption = "&Save"
esc = 1
Open userPath For Output As #1
a = a + 1
Print #1, LTrim(a)
Close #1
txtUserID.Text = a
Else
 If txtUserName.Text <> "" And txtName.Text <> "" _
 And txtPosition.Text <> "" And cboAccLevel.Text <> "" _
 And txtPassword.Text <> "" Then
 Call ConFields
 rsUser.Update
 MsgBox "New user has been added.", vbInformation, Me.Caption
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
If Err.Number = 3022 Then
MsgBox "User Name: " & txtUserName.Text & " already exists.", vbInformation, Me.Caption
Else
MsgBox Err.Description, vbInformation, Me.Caption
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
If txtUserName.Text <> "" Then
Set rsUser = db.OpenRecordset("SELECT * FROM Users WHERE UserName = '" & txtUserName.Text & "'")
reply = MsgBox("User Name: " & rsUser.Fields!UserName & Chr(13) & "Password: " & String(Len(rsUser.Fields!Password), "*") & Chr(13) & "Are you sure you want to Delete this user?", vbQuestion + vbYesNo, Me.Caption)
If reply = vbYes Then
rsUser.Delete
Call clrTxt
MsgBox "User deleted.", vbInformation, Me.Caption
cmdUpdate.Caption = "&Search"
cmdAdd.Enabled = True
Frame1.Enabled = False
Set rsUser = db.OpenRecordset("SELECT * FROM Users")
End If
Else
MsgBox "No user to delete.", vbInformation, Me.Caption
End If
End Sub

Private Sub cmdUpdate_Click()
If cmdUpdate.Caption = "&Search" Then
entry = InputBox("Enter User Name: ", Me.Caption)
rsUser.FindFirst "UserName = '" & entry & "'"
If rsUser.NoMatch = False Then
Call RetFields
cmdUpdate.Caption = "&Update"
Frame1.Enabled = True
cmdAdd.Enabled = False
tempUser = rsUser.Fields!UserName
Else
MsgBox "User Name: " & entry & " not found!", vbInformation, Me.Caption
End If
Else
 If txtUserName.Text <> "" And txtName.Text <> "" _
 And txtPosition.Text <> "" And cboAccLevel.Text <> "" _
 And txtPassword.Text <> "" Then
 If tempUser = txtUserName.Text Then
Call updateRec
MsgBox "User updated", vbInformation, Me.Caption
 Set rsUser = db.OpenRecordset("SELECT * FROM Users")
Frame1.Enabled = False
cmdAdd.Enabled = True
cmdUpdate.Caption = "&Search"
Call clrTxt
 Exit Sub
 End If
 If tempUser <> txtUserName.Text Then
 Set rsUser = db.OpenRecordset("SELECT * FROM Users WHERE UserName = '" & txtUserName.Text & "'")
  If rsUser.BOF = True Then
Call updateRec
MsgBox "User updated", vbInformation, Me.Caption
 Set rsUser = db.OpenRecordset("SELECT * FROM Users")
Frame1.Enabled = False
cmdAdd.Enabled = True
cmdUpdate.Caption = "&Search"
Call clrTxt
 Else
 MsgBox "User Name: " & txtUserName.Text & " already exists!", vbInformation, Me.Caption
 txtUserName.SetFocus
 SendKeys "{home}+{end}"
 Exit Sub
 End If
 Exit Sub
 End If
 Else
 MsgBox "Fill up all the required fields.", vbInformation, Me.Caption
 Exit Sub
 End If
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
            rsUser.CancelUpdate
            Call clrTxt
            cmdAdd.Caption = "&Add"
            Frame1.Enabled = False
            cmdDelete.Enabled = True
            cmdUpdate.Enabled = True
            Open userPath For Output As #1
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
Set rsUser = db.OpenRecordset("SELECT * FROM Users")
esc = 0
userPath = App.Path & "\Users"
Open userPath For Input As #1
If Not EOF(1) Then
Line Input #1, b
a = Val(b)
End If
Close #1
End Sub

Private Sub clrTxt()
For Each txt In Me.Controls
If TypeOf txt Is TextBox Then
txt.Text = ""
End If
Next
cboAccLevel.ListIndex = -1
End Sub

Private Sub ConFields()
rsUser.Fields!UserID = txtUserID.Text
rsUser.Fields!UserName = txtUserName.Text
rsUser.Fields!UserFullName = txtName.Text
rsUser.Fields!UserPosition = txtPosition.Text
rsUser.Fields!AccLevel = cboAccLevel.Text
rsUser.Fields!Password = txtPassword.Text
End Sub

Private Sub RetFields()
txtUserID.Text = rsUser.Fields!UserID
txtUserName.Text = rsUser.Fields!UserName
txtName.Text = rsUser.Fields!UserFullName
txtPosition.Text = rsUser.Fields!UserPosition
cboAccLevel.Text = rsUser.Fields!AccLevel
txtPassword.Text = rsUser.Fields!Password
End Sub

Private Sub updateRec()
sqlStr = "UPDATE Users SET " _
& " UserID = '" & txtUserID.Text & "'" _
& ", UserName = '" & txtUserName.Text & "'" _
& ", UserFullName = '" & txtName.Text & "'" _
& ", UserPosition = '" & txtPosition.Text & "'" _
& ", AccLevel = '" & cboAccLevel.Text & "'" _
& ", Password = '" & txtPassword.Text & "' WHERE UserID = '" & txtUserID.Text & "'"
execQuery (sqlStr)
End Sub

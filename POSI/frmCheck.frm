VERSION 5.00
Begin VB.Form frmCheck 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment by Check"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
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
   ScaleHeight     =   2145
   ScaleWidth      =   5205
   Begin VB.Frame Frame1 
      Height          =   2040
      Left            =   45
      TabIndex        =   3
      Top             =   45
      Width           =   5100
      Begin VB.TextBox txtBank 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1485
         TabIndex        =   2
         Top             =   1080
         Width           =   3435
      End
      Begin VB.TextBox txtCheckNumber 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1485
         TabIndex        =   1
         Top             =   675
         Width           =   3435
      End
      Begin VB.TextBox txtCheckAmount 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1485
         TabIndex        =   0
         Top             =   270
         Width           =   3435
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   375
         Left            =   3690
         TabIndex        =   4
         Top             =   1530
         Width           =   1275
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "<Esc>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2340
         TabIndex        =   9
         Top             =   1710
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Press              to Cancel."
         Height          =   195
         Left            =   1890
         TabIndex        =   8
         Top             =   1710
         Width           =   1755
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Bank:"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   1125
         Width           =   405
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   1530
         Top             =   1125
         Width           =   3435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Check Number:"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   1530
         Top             =   720
         Width           =   3435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Check Amount:"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   315
         Width           =   1095
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   1530
         Top             =   315
         Width           =   3435
      End
   End
End
Attribute VB_Name = "frmCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Const MF_BYPOSITION = &H400&

Private Sub RemoveMenus()
    Dim hMenu As Long
    hMenu = GetSystemMenu(hWnd, False)
    DeleteMenu hMenu, 6, MF_BYPOSITION
End Sub

Private Sub cmdOK_Click()
frmPOSI.PType = 1
frmPOSI.CAmount = txtCheckAmount.Text
frmPOSI.CNumber = txtCheckNumber.Text
frmPOSI.Banko = txtBank.Text
frmPOSI.lblAmount.Caption = "Check Amount:"
frmPOSI.txtCash.Text = CAmount
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
        Unload Me
End Select
End Sub

Private Sub Form_Load()
RemoveMenus
Me.Top = (Screen.Height / 2) - (Me.Height / 2)
Me.Left = (Screen.Width / 2) - (Me.Width / 2)
End Sub

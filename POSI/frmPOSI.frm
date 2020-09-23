VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPOSI 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11865
   ControlBox      =   0   'False
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
   ScaleHeight     =   7650
   ScaleWidth      =   11865
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab1 
      Height          =   2535
      Left            =   3375
      TabIndex        =   52
      Top             =   2565
      Visible         =   0   'False
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   4471
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Payment By Check"
      TabPicture(0)   =   "frmPOSI.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame10"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame10 
         Height          =   2040
         Left            =   135
         TabIndex        =   53
         Top             =   360
         Width           =   5100
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   3690
            TabIndex        =   13
            Top             =   1530
            Width           =   1275
         End
         Begin VB.TextBox txtBank 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1485
            TabIndex        =   11
            Top             =   1080
            Width           =   3435
         End
         Begin VB.TextBox txtCheckNumber 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1485
            TabIndex        =   10
            Top             =   675
            Width           =   3435
         End
         Begin VB.TextBox txtCheckAmount 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1485
            TabIndex        =   9
            Top             =   270
            Width           =   3435
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "&OK"
            Height          =   375
            Left            =   2385
            TabIndex        =   12
            Top             =   1530
            Width           =   1275
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Bank:"
            Height          =   195
            Left            =   180
            TabIndex        =   56
            Top             =   1125
            Width           =   405
         End
         Begin VB.Shape Shape15 
            BackColor       =   &H00000000&
            BackStyle       =   1  'Opaque
            Height          =   285
            Left            =   1530
            Top             =   1125
            Width           =   3435
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Check Number:"
            Height          =   195
            Left            =   180
            TabIndex        =   55
            Top             =   720
            Width           =   1095
         End
         Begin VB.Shape Shape16 
            BackColor       =   &H00000000&
            BackStyle       =   1  'Opaque
            Height          =   285
            Left            =   1530
            Top             =   720
            Width           =   3435
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Check Amount:"
            Height          =   195
            Left            =   180
            TabIndex        =   54
            Top             =   315
            Width           =   1095
         End
         Begin VB.Shape Shape17 
            BackColor       =   &H00000000&
            BackStyle       =   1  'Opaque
            Height          =   285
            Left            =   1530
            Top             =   315
            Width           =   3435
         End
      End
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   3015
      Sorted          =   -1  'True
      TabIndex        =   51
      Top             =   4005
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.Frame Frame8 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5100
      Left            =   7020
      TabIndex        =   34
      Top             =   1530
      Width           =   4740
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   3105
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Frame Frame9 
         Caption         =   "Current User:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1770
         Left            =   135
         TabIndex        =   43
         Top             =   1260
         Width           =   4470
         Begin VB.TextBox txtUserFullName 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   180
            Locked          =   -1  'True
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   585
            Width           =   4065
         End
         Begin VB.TextBox txtPosition 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   180
            Locked          =   -1  'True
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   1260
            Width           =   4065
         End
         Begin VB.Shape Shape13 
            BackColor       =   &H00000000&
            BackStyle       =   1  'Opaque
            Height          =   285
            Left            =   225
            Top             =   630
            Width           =   4065
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Name:"
            Height          =   195
            Left            =   180
            TabIndex        =   47
            Top             =   315
            Width           =   465
         End
         Begin VB.Shape Shape12 
            BackColor       =   &H00000000&
            BackStyle       =   1  'Opaque
            Height          =   285
            Left            =   225
            Top             =   1305
            Width           =   4065
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Position:"
            Height          =   195
            Left            =   180
            TabIndex        =   46
            Top             =   990
            Width           =   615
         End
      End
      Begin VB.TextBox txtORNumber 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1395
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   675
         Width           =   3210
      End
      Begin VB.TextBox txtInvoiceNumber 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1395
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   315
         Width           =   3210
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   2385
         Picture         =   "frmPOSI.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdLogOff 
         Caption         =   "&Log Off"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   3555
         Picture         =   "frmPOSI.frx":0420
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   3465
         Width           =   1095
      End
      Begin VB.Line Line1 
         X1              =   135
         X2              =   4635
         Y1              =   1215
         Y2              =   1215
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         Caption         =   "Email:  j_villaflor@yahoo.com"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2430
         TabIndex        =   42
         Top             =   4725
         Width           =   2100
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Program Created By:  Jay E. Villaflor"
         Height          =   195
         Left            =   1935
         TabIndex        =   41
         Top             =   4455
         Width           =   2610
      End
      Begin VB.Label Label10 
         Caption         =   "OR Number:"
         Height          =   285
         Left            =   135
         TabIndex        =   39
         Top             =   720
         Width           =   1185
      End
      Begin VB.Shape Shape11 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   1440
         Top             =   720
         Width           =   3210
      End
      Begin VB.Label Label9 
         Caption         =   "Invoice Number:"
         Height          =   285
         Left            =   135
         TabIndex        =   37
         Top             =   360
         Width           =   1275
      End
      Begin VB.Shape Shape10 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   1440
         Top             =   360
         Width           =   3210
      End
   End
   Begin VB.Frame Frame7 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   7020
      TabIndex        =   33
      Top             =   0
      Width           =   4740
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Point On-Sale and Inventory System"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   585
         TabIndex        =   40
         Top             =   405
         Width           =   3525
      End
   End
   Begin VB.Frame Frame6 
      Enabled         =   0   'False
      Height          =   915
      Left            =   7020
      TabIndex        =   30
      Top             =   6660
      Width           =   4740
      Begin VB.TextBox txtChange 
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
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   315
         Width           =   3570
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   1080
         Top             =   360
         Width           =   3570
      End
      Begin VB.Label Label8 
         Caption         =   "Change:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   32
         Top             =   360
         Width           =   780
      End
   End
   Begin VB.Frame Frame5 
      Enabled         =   0   'False
      Height          =   690
      Left            =   90
      TabIndex        =   27
      Top             =   6885
      Width           =   6900
      Begin VB.TextBox txtCash 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   1755
         TabIndex        =   7
         Top             =   270
         Width           =   3795
      End
      Begin VB.CommandButton cmdPayCheck 
         Caption         =   "&Pay Check"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5670
         TabIndex        =   28
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label lblAmount 
         Caption         =   "Cash Amount:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   29
         Top             =   270
         Width           =   1590
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   1800
         Top             =   315
         Width           =   3795
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Checked by:    [Enter User ID or Select User:]   "
      Enabled         =   0   'False
      Height          =   690
      Left            =   90
      TabIndex        =   24
      Top             =   6165
      Width           =   6900
      Begin VB.ComboBox cboID 
         Height          =   315
         Left            =   90
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   270
         Width           =   1365
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5445
         TabIndex        =   26
         Top             =   270
         Width           =   1320
      End
      Begin VB.TextBox txtCheckedBy 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   270
         Width           =   3705
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   1575
         Top             =   315
         Width           =   3705
      End
   End
   Begin VB.Frame Frame3 
      Enabled         =   0   'False
      Height          =   735
      Left            =   90
      TabIndex        =   21
      Top             =   5400
      Width           =   6900
      Begin VB.TextBox txtTotalAmount 
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   1620
         TabIndex        =   22
         Top             =   225
         Width           =   5100
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   330
         Left            =   1665
         Top             =   270
         Width           =   5100
      End
      Begin VB.Label Label6 
         Caption         =   "Total Amount:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   23
         Top             =   270
         Width           =   1860
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Item(s)"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3840
      Left            =   90
      TabIndex        =   19
      Top             =   1530
      Width           =   6900
      Begin MSComctlLib.ListView lv1 
         Height          =   3300
         Left            =   90
         TabIndex        =   20
         ToolTipText     =   "Double Click to edit.  Press <Delete> to Delete."
         Top             =   315
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   5821
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Product Code"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Qty"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Price"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Disc"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Subtotal"
            Object.Width           =   2152
         EndProperty
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   3300
         Left            =   180
         Top             =   405
         Width           =   6630
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Item Entry"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   6900
      Begin VB.TextBox txtSupply 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   315
         Width           =   960
      End
      Begin VB.CommandButton cmdAddItem 
         Caption         =   "&Add Item"
         Enabled         =   0   'False
         Height          =   330
         Left            =   5580
         TabIndex        =   5
         Top             =   1035
         Width           =   1185
      End
      Begin VB.TextBox txtQuantity 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3870
         MaxLength       =   5
         TabIndex        =   4
         Top             =   1035
         Width           =   1410
      End
      Begin VB.TextBox txtDiscount 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1350
         MaxLength       =   2
         TabIndex        =   3
         Top             =   1035
         Width           =   510
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   675
         Width           =   5370
      End
      Begin VB.TextBox txtProdCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1350
         MaxLength       =   10
         TabIndex        =   2
         Top             =   315
         Width           =   1545
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "[F2]  -  Find Product"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3060
         TabIndex        =   57
         Top             =   360
         Width           =   1440
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Available:"
         Height          =   195
         Left            =   4995
         TabIndex        =   50
         Top             =   360
         Width           =   705
      End
      Begin VB.Shape Shape14 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   5805
         Top             =   360
         Width           =   960
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   3915
         Top             =   1080
         Width           =   1410
      End
      Begin VB.Label Label5 
         Caption         =   "Quantity:"
         Height          =   285
         Left            =   3060
         TabIndex        =   18
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1980
         TabIndex        =   17
         Top             =   1080
         Width           =   240
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   1395
         Top             =   1080
         Width           =   510
      End
      Begin VB.Label Label3 
         Caption         =   "Discount:"
         Height          =   285
         Left            =   180
         TabIndex        =   16
         Top             =   1080
         Width           =   1050
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   1395
         Top             =   720
         Width           =   5370
      End
      Begin VB.Label Label2 
         Caption         =   "Description:"
         Height          =   285
         Left            =   180
         TabIndex        =   14
         Top             =   720
         Width           =   1050
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   1395
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Product Code:"
         Height          =   285
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Width           =   1050
      End
   End
End
Attribute VB_Name = "frmPOSI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim invoicePath As String
Dim salesPath As String
Dim s1 As Integer
Dim s2
Dim a As Integer
Dim b
Dim temp
Dim rsUser As DAO.Recordset
Dim rsUserLog As DAO.Recordset
Dim sqlStr As String
Dim rsInvoice As DAO.Recordset
Dim rsProducts As DAO.Recordset
Dim rsSalesDetails As DAO.Recordset
Dim rsSales As DAO.Recordset
Dim comp As Double
Dim tempcom As Double
Dim valTemp As Double
Public PType As Byte
Public CAmount
Public CNumber
Public Banko
Public tempID
Dim qtyTemp As Integer

Private Sub cboID_Change()
If cboID.Text <> "" Then
Set rsUser = db.OpenRecordset("SELECT * FROM Users WHERE UserID = '" & LTrim(cboID.Text) & "'")
If rsUser.RecordCount <> 0 Then
txtCheckedBy.Text = rsUser.Fields!UserFullName
Else
txtCheckedBy.Text = ""
End If
Else
txtCheckedBy.Text = ""
End If
End Sub

Private Sub cboID_Click()
If cboID.Text <> "" Then
Set rsUser = db.OpenRecordset("SELECT * FROM Users WHERE UserID = '" & LTrim(cboID.Text) & "'")
If rsUser.RecordCount <> 0 Then
txtCheckedBy.Text = rsUser.Fields!UserFullName
Else
txtCheckedBy.Text = ""
End If
Else
txtCheckedBy.Text = ""
End If
End Sub

Private Sub cmdAddItem_Click()
Set itmFound = lv1.FindItem(txtProdCode.Text)
If itmFound Is Nothing Then
valTemp = 0
If txtDiscount.Text = "" Then
txtDiscount.Text = 0
comp = (Val(txtQuantity.Text) * Val(rsProducts.Fields!CostPrice))
Else
tempcomp = (Val(txtQuantity.Text) * Val(rsProducts.Fields!CostPrice)) * (Val(txtDiscount.Text) / 100)
comp = (Val(txtQuantity.Text) * Val(rsProducts.Fields!CostPrice)) - Val(tempcomp)
End If
Set rsProducts = db.OpenRecordset("SELECT * FROM Products WHERE ProdCode = '" & txtProdCode.Text & "'")
rsProducts.Edit
rsProducts.Fields!Qty = (Val(rsProducts.Fields!Qty) + Val(qtyTemp)) - Val(txtQuantity.Text)
rsProducts.Update
Set j = lv1.ListItems.Add(, , txtProdCode.Text)
j.SubItems(1) = txtDescription.Text
j.SubItems(2) = txtQuantity.Text
j.SubItems(3) = rsProducts.Fields!CostPrice
j.SubItems(4) = txtDiscount.Text
j.SubItems(5) = comp
txtProdCode.Text = ""
txtDiscount.Text = ""
txtQuantity.Text = ""
txtProdCode.SetFocus
For i = 1 To lv1.ListItems.Count
valTemp = Val(valTemp) + Val(lv1.ListItems.Item(i).SubItems(5))
txtTotalAmount.Text = valTemp
Next
Else
MsgBox "Product code already exists in the list.", vbInformation, mdiPOSI.Caption
txtProdCode.Text = ""
txtDiscount.Text = ""
txtQuantity.Text = ""
End If
End Sub

Private Sub cmdCancel_Click()
tab1.Visible = False
End Sub

Private Sub cmdLogOff_Click()
If cmdNew.Caption = "&Save" Then
MsgBox "Save first your transaction.", vbInformation, mdiPOSI.Caption
Exit Sub
End If
reply = MsgBox("Do you want to log off now?", vbQuestion + vbYesNo, mdiPOSI.Caption)
If reply = vbYes Then
    Set rsUserLog = db.OpenRecordset("SELECT * FROM Login WHERE LogInNo = " & Text1.Text & "")
    rsUserLog.Edit
    rsUserLog.Fields!LogOut = Format(Time, "hh:mm:ss AM/PM")
    rsUserLog.Update
End
End If
End Sub

Private Sub cmdNew_Click()
If cmdNew.Caption = "&New" Then
Open invoicePath For Output As #1
a = a + 1
Print #1, LTrim(a)
Close #1
txtInvoiceNumber.Text = Format(LTrim(a), "#000000#")
Frame1.Enabled = True
Frame3.Enabled = True
Frame4.Enabled = True
Frame5.Enabled = True
txtProdCode.SetFocus
cmdNew.Caption = "&Save"
lblAmount.Caption = "Cash Amount:"
'-----------------------------------------------------
Else
If txtCheckedBy.Text = "" Then
MsgBox "Please enter name of checker.", vbInformation, mdiPOSI.Caption
cboID.SetFocus
Exit Sub
End If
If txtCash.Text = "" Then
MsgBox "Please pay first the amount of Php " & txtTotalAmount.Text, vbInformation, mdiPOSI.Caption
txtCash.SetFocus
Exit Sub
End If
If txtChange.Text = "" Then
Call txtCash_KeyDown(13, 0)
End If
Frame1.Enabled = False
Frame3.Enabled = False
Frame4.Enabled = False
Frame5.Enabled = False
cmdNew.Caption = "&New"
mdiPOSI.st1.Panels(5).Text = Format(Val(mdiPOSI.st1.Panels(5).Text) + Val(txtTotalAmount.Text), "###,###,###.#0")
rsInvoice.AddNew
rsInvoice.Fields!InvoiceID = txtInvoiceNumber.Text
rsInvoice.Fields!OrNo = txtORNumber.Text
rsInvoice.Fields!PaymentType = PType
rsInvoice.Fields!DateSold = Date
rsInvoice.Fields!TimeSold = Format(Time, "hh:mm:ss AM/PM")
rsInvoice.Fields!CashAmount = txtCash.Text
rsInvoice.Fields!CheckAmount = Val(CAmount)
rsInvoice.Fields!CheckNo = Val(CNumber)
rsInvoice.Fields!Bank = Banko
rsInvoice.Update
For i = 1 To lv1.ListItems.Count
rsSalesDetails.AddNew
rsSalesDetails.Fields!ProdCode = lv1.ListItems.Item(i).Text
rsSalesDetails.Fields!QtySold = Val(lv1.ListItems.Item(i).SubItems(2))
rsSalesDetails.Fields!DateSold = Date
rsSalesDetails.Fields!TimeSold = Format(Time, "hh:mm:ss AM/PM")
rsSalesDetails.Fields!CheckedBy = cboID.Text
rsSalesDetails.Fields!Cashier = temp
rsSalesDetails.Fields!InvoiceID = txtInvoiceNumber.Text
rsSalesDetails.Fields!Discount = Val(lv1.ListItems.Item(i).SubItems(4))
rsSalesDetails.Update
Next
Set rsSales = db.OpenRecordset("SELECT * FROM Sales WHERE SalesID = '" & s1 & "'")
If rsSales.BOF = True Then
rsSales.AddNew
Call salesCon
rsSales.Update
Else
rsSales.Edit
Call salesCon
rsSales.Update
End If
MsgBox "Your change is : Php " & txtChange.Text & Chr(13) & "Transaction Saved.", vbInformation, mdiPOSI.Caption
End If
lv1.ListItems.Clear
cboID.Text = ""
txtCash.Text = ""
txtTotalAmount.Text = ""
txtChange.Text = ""
lblAmount.Caption = "Cash Amount:"
End Sub

Private Sub cmdOK_Click()
If txtCheckAmount.Text = "" Then
MsgBox "Please enter Check Amount.", vbInformation, mdiPOSI.Caption
txtCheckAmount.SetFocus
Exit Sub
End If
If txtCheckNumber.Text = "" Then
MsgBox "Please enter Check Number.", vbInformation, mdiPOSI.Caption
txtCheckNumber.SetFocus
Exit Sub
End If
If txtBank.Text = "" Then
MsgBox "Please enter Name of Bank.", vbInformation, mdiPOSI.Caption
txtBank.SetFocus
Exit Sub
End If
If IsNumeric(txtCheckAmount.Text) = False Then
MsgBox "Check Amount must be in numeric type.", vbInformation, mdiPOSI.Caption
txtCheckAmount.SetFocus
Exit Sub
End If
PType = 1
CAmount = txtCheckAmount.Text
CNumber = txtCheckNumber.Text
Banko = txtBank.Text
lblAmount.Caption = "Check Amount:"
txtCash.Text = CAmount
tab1.Visible = False
Call txtCash_KeyDown(13, 0)
End Sub

Private Sub cmdPayCheck_Click()
tab1.Visible = True
txtCheckAmount.SetFocus
End Sub

Private Sub cmdSelect_Click()
List1.Visible = True
List1.SetFocus
End Sub

Private Sub Form_Activate()
Set rsUser = db.OpenRecordset("SELECT * FROM Users")
List1.Clear
Do While Not rsUser.EOF
List1.AddItem rsUser.Fields!UserFullName
rsUser.MoveNext
Loop
Set rsUser = db.OpenRecordset("SELECT * FROM Users")
cboID.Clear
Do While Not rsUser.EOF
cboID.AddItem rsUser.Fields!UserID
rsUser.MoveNext
Loop
cboID.Text = tempID
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
    Dim reply
    reply = MsgBox("Do you want to cancel purchase?", vbQuestion + vbYesNo, mdiPOSI.Caption)
    If reply = vbYes Then
        lv1.ListItems.Clear
        txtProdCode.Text = ""
        txtTotalAmount.Text = ""
        txtCheckedBy.Text = ""
        cboID.Text = ""
        txtCash.Text = ""
        txtChange.Text = ""
        tab1.Visible = False
        List1.Visible = False
        txtCheckAmount.Text = ""
        txtBank.Text = ""
        txtCheckNumber.Text = ""
        txtDiscount.Text = ""
        txtQuantity.Text = ""
        cmdNew.Caption = "&New"
    End If
    Case 13
    SendKeys "{tab}"
End Select
End Sub

Private Sub Form_Load()
PType = 0
CAmount = ""
CNumber = ""
Banko = " "
Me.Top = 22
Me.Left = 22
openDB
Set rsSalesDetails = db.OpenRecordset("SELECT * FROM SalesDetails")
Set rsInvoice = db.OpenRecordset("SELECT * FROM Invoices")
invoicePath = App.Path & "\invoice"
Open invoicePath For Input As #1
If Not EOF(1) Then
Line Input #1, b
a = Val(b)
End If
Close #1
userLog = App.Path & "\userLogFile"
Open userLog For Input As #1
If Not EOF(1) Then
Line Input #1, temp
End If
Close #1
salesPath = App.Path & "\Sales"
Open salesPath For Input As #1
If Not EOF(1) Then
Line Input #1, s2
s1 = Val(s2)
End If
Close #1
Open salesPath For Output As #1
s1 = s1 + 1
Print #1, LTrim(s1)
Close #1
Set rsUser = db.OpenRecordset("SELECT * FROM Users WHERE UserID = '" & temp & "'")
txtUserFullName.Text = rsUser.Fields!UserFullName
txtPosition.Text = rsUser.Fields!UserPosition
End Sub

Private Sub List1_DblClick()
row = List1.ListIndex
Set rsUser = db.OpenRecordset("SELECT * FROM Users WHERE UserFullName = '" & List1.List(row) & "'")
txtCheckedBy.Text = List1.List(row)
cboID.Text = rsUser.Fields!UserID
List1.Visible = False
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 13
        row = List1.ListIndex
        Set rsUser = db.OpenRecordset("SELECT * FROM Users WHERE UserFullName = '" & List1.List(row) & "'")
        txtCheckedBy.Text = List1.List(row)
        cboID.Text = rsUser.Fields!UserID
        List1.Visible = False
End Select
End Sub

Private Sub lv1_DblClick()
rw = lv1.SelectedItem.Index
txtProdCode.Text = lv1.ListItems.Item(rw).Text
txtDiscount.Text = lv1.ListItems.Item(rw).SubItems(4)
txtQuantity.Text = lv1.ListItems.Item(rw).SubItems(2)
qtyTemp = txtQuantity.Text
lv1.ListItems.Remove (rw)
valTemp = 0
For i = 1 To lv1.ListItems.Count
valTemp = Val(valTemp) + Val(lv1.ListItems.Item(i).SubItems(5))
txtTotalAmount.Text = valTemp
Next
End Sub

Private Sub lv1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyDelete
    If lv1.ListItems.Count <> 0 Then
rw = lv1.SelectedItem.Index
Set rsProducts = db.OpenRecordset("SELECT * FROM Products WHERE ProdCode = '" & lv1.ListItems.Item(rw).Text & "'")
rsProducts.Edit
rsProducts.Fields!Qty = Val(rsProducts.Fields!Qty) + Val(lv1.ListItems.Item(rw).SubItems(2))
rsProducts.Update
lv1.ListItems.Remove (rw)
valTemp = 0
For i = 1 To lv1.ListItems.Count
valTemp = Val(valTemp) + Val(lv1.ListItems.Item(i).SubItems(5))
txtTotalAmount.Text = valTemp
Next
End If
End Select
End Sub

Private Sub txtCash_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 13
        txtChange.Text = Val(txtCash.Text) - Val(txtTotalAmount.Text)
        If txtChange.Text < 0 Then
            MsgBox "Cash is smaller than the Total Amount.", vbInformation, mdiPOSI.Caption
            txtCash.SetFocus
            SendKeys "{home}+{end}"
            txtChange.Text = ""
        End If
End Select
End Sub

Private Sub txtCheckNumber_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
If KeyAscii = vbKeyBack Then
KeyAscii = vbKeyBack
Else
KeyAscii = 0
End If
End If
End Sub

Private Sub txtDescription_Change()
test
End Sub

Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
If KeyAscii = vbKeyBack Then
KeyAscii = vbKeyBack
Else
KeyAscii = 0
End If
End If
End Sub

Private Sub txtInvoiceNumber_Change()
txtORNumber.Text = txtInvoiceNumber.Text
End Sub


Private Sub txtProdCode_Change()
On Error Resume Next
test
If txtProdCode.Text <> "" Then
Set rsProducts = db.OpenRecordset("SELECT * FROM Products WHERE ProdCode = '" & txtProdCode.Text & "'")
If rsProducts.RecordCount <> 0 Then
txtSupply.Text = rsProducts.Fields!Qty
txtDescription.Text = rsProducts.Fields!ProdDescription
Else
txtSupply.Text = ""
txtDescription.Text = ""
End If
Else
txtSupply.Text = ""
txtDescription.Text = ""
End If
End Sub

Private Sub test()
If txtProdCode.Text <> "" And txtQuantity.Text <> "" And txtDescription.Text <> "" Then
cmdAddItem.Enabled = True
Else
cmdAddItem.Enabled = False
End If
End Sub

Private Sub txtProdCode_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF2
        frmViewer.Show
End Select
End Sub

Private Sub txtQuantity_Change()
test
If txtQuantity.Text <> "" Then
If Val(txtSupply.Text) <= 0 Then
MsgBox "Insufficient supply.", vbInformation, mdiPOSI.Caption
End If
End If
End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
If KeyAscii = vbKeyBack Then
KeyAscii = vbKeyBack
Else
KeyAscii = 0
End If
End If
End Sub

Private Sub salesCon()
rsSales.Fields!SalesID = s1
rsSales.Fields!TotalSales = Val(mdiPOSI.st1.Panels(5).Text)
rsSales.Fields!TimeGenerated = Format(Time, "hh:mm:ss AM/PM")
rsSales.Fields!DateGenerated = Date
rsSales.Fields!UserID = temp
End Sub

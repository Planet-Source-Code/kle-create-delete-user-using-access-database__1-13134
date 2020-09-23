VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Main Menu"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   5175
      Begin VB.PictureBox Picture2 
         Height          =   615
         Left            =   120
         ScaleHeight     =   555
         ScaleWidth      =   4800
         TabIndex        =   23
         Top             =   2160
         Width           =   4860
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add Record"
            Height          =   375
            Left            =   120
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete Record"
            Height          =   375
            Left            =   1680
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "Next Record"
            Height          =   375
            Left            =   3240
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.TextBox txtFirstName 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtLastName 
         Height          =   285
         Left            =   1800
         TabIndex        =   12
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox cboSex 
         Height          =   315
         Left            =   3480
         TabIndex        =   11
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtTelephone 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   1800
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtCity 
         Height          =   285
         Left            =   3480
         TabIndex        =   8
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtCountry 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtBirthDate 
         Height          =   285
         Left            =   3480
         TabIndex        =   5
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lblFirstName 
         AutoSize        =   -1  'True
         Caption         =   "First Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   795
      End
      Begin VB.Label lblLastName 
         AutoSize        =   -1  'True
         Caption         =   "Last Name:"
         Height          =   195
         Left            =   1800
         TabIndex        =   21
         Top             =   240
         Width           =   810
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         Caption         =   "Sex:"
         Height          =   195
         Left            =   3480
         TabIndex        =   20
         Top             =   240
         Width           =   315
      End
      Begin VB.Label lblTelephone 
         AutoSize        =   -1  'True
         Caption         =   "Telephone:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   810
      End
      Begin VB.Label lblAddress 
         AutoSize        =   -1  'True
         Caption         =   "Address:"
         Height          =   195
         Left            =   1800
         TabIndex        =   18
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblCity 
         AutoSize        =   -1  'True
         Caption         =   "City:"
         Height          =   195
         Left            =   3480
         TabIndex        =   17
         Top             =   840
         Width           =   300
      End
      Begin VB.Label lblCountry 
         AutoSize        =   -1  'True
         Caption         =   "Country:"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   585
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         Caption         =   "Email:"
         Height          =   195
         Left            =   1800
         TabIndex        =   15
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label lblBirthDate 
         AutoSize        =   -1  'True
         Caption         =   "Birth Date:"
         Height          =   195
         Left            =   3480
         TabIndex        =   14
         Top             =   1440
         Width           =   750
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   5340
      TabIndex        =   0
      Top             =   0
      Width           =   5340
      Begin VB.CommandButton cmdChangePWD 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Change PWD"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdUserProfile 
         BackColor       =   &H00C0FFFF&
         Caption         =   "User Profile"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   1335
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C000&
         X1              =   120
         X2              =   120
         Y1              =   -120
         Y2              =   840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C000&
         X1              =   240
         X2              =   240
         Y1              =   -120
         Y2              =   840
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdChangePWD_Click()
    frmMain.Visible = False
    frmUserProfile.Show
    frmUserProfile!cmdChangePWD.Caption = "Change PWD"
    frmUserProfile!cmdAddUser.Enabled = False
    frmUserProfile!cmdDeleteUser.Enabled = False
    frmUserProfile!cmdFindUser.Enabled = False
    frmUserProfile!txtPassword.Enabled = True
    frmUserProfile!cmdChangePWD.Enabled = True
    frmUserProfile!cmdClose.Enabled = False
    frmUserProfile!tmrFindUser.Enabled = True
    frmUserProfile!txtPassword.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdUserProfile_Click()
    frmMain.Visible = False
    frmUserProfile.Show
End Sub

Private Sub Form_Load()
    frmMain.Height = 4200
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmTemp
End Sub


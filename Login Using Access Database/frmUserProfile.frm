VERSION 5.00
Begin VB.Form frmUserProfile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Profile"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrFindUser 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1200
      Top             =   1680
   End
   Begin VB.CommandButton cmdFindUser1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Find User"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   5130
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   5130
      Begin VB.CommandButton cmdDeleteUser 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Delete User"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdChangePwd 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Change User PWD/PRF"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdFindUser 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Find User"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdAddUser 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Add User"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Add USer"
         Top             =   120
         Width           =   855
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C000&
         X1              =   240
         X2              =   240
         Y1              =   -120
         Y2              =   840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C000&
         X1              =   120
         X2              =   120
         Y1              =   -120
         Y2              =   840
      End
   End
   Begin VB.ComboBox cboLevel 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3480
      TabIndex        =   5
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtPassword 
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtUsername 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblLevel 
      AutoSize        =   -1  'True
      Caption         =   "Level:"
      Height          =   195
      Left            =   3480
      TabIndex        =   4
      Top             =   960
      Width           =   435
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblUsername 
      AutoSize        =   -1  'True
      Caption         =   "Username:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   765
   End
End
Attribute VB_Name = "frmUserProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dbWorkspace As Workspace
Private dbDatabase As Database
Private dbTable As Recordset
Public KLEDll As New prjEncPwd.clsEncPwd

Private Sub cboLevel_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub cmdAddUser_Click()
    If cmdAddUser.Caption = "Add User" Then
    cmdAddUser.Caption = "Save"
    txtUserName.Enabled = True
    txtPassword.Enabled = True
    cboLevel.Enabled = True
    cmdDeleteUser.Enabled = False
    cmdChangePwd.Enabled = False
    cmdFindUser.Enabled = False
    cmdClose.Enabled = False
    txtUserName.Text = ""
    txtPassword.Text = ""
    cboLevel.Text = "User"
    txtUserName.SetFocus
    Exit Sub
    End If
    '----------
    If txtUserName = "" Or txtPassword = "" Then
    Exit Sub
    End If
    '----------
    dbTable.MoveLast
    dbTable.AddNew
    '**********
    txtPassword.Text = KLEDll.Encrypt(txtPassword.Text)
    '**********
    dbTable!UserName = txtUserName.Text
    dbTable!Password = txtPassword.Text
    dbTable!Level = cboLevel.Text
    dbTable.Update
    MsgBox "User ***" + txtUserName.Text + "*** created succesfully! (Level: ***" + cboLevel.Text + "***)", vbOKOnly + vbCritical + vbSystemModal, "Phone Book - [User Profile]"
    txtUserName.Text = ""
    txtPassword.Text = ""
    cboLevel.ListIndex = 0
    '----------
    If cmdAddUser.Caption = "Save" Then
    cmdAddUser.Caption = "Add User"
    txtUserName.Enabled = False
    txtPassword.Enabled = False
    cboLevel.Enabled = False
    cmdDeleteUser.Enabled = False
    cmdFindUser.Enabled = True
    cmdClose.Enabled = True
    End If
End Sub

Private Sub cmdChangePWD_Click()
    If txtPassword.Text = "" Then Exit Sub
    If cmdChangePwd.Caption = "Change PWD" Then
    dbTable.Edit
    '**********
    If txtUserName.Text = "ADMIN" And txtPassword.Text = "admin" Then
    dbTable!UserName = txtUserName.Text
    dbTable!Password = txtPassword.Text
    dbTable!Level = cboLevel.Text
    dbTable.Update
    cmdChangePwd.Enabled = False
    cmdClose.Enabled = True
    txtPassword.Enabled = False
    Exit Sub
    Else
    txtPassword.Text = KLEDll.Encrypt(txtPassword.Text)
    End If
    '**********
    dbTable!UserName = txtUserName.Text
    dbTable!Password = txtPassword.Text
    dbTable!Level = cboLevel.Text
    dbTable.Update
    cmdChangePwd.Enabled = True
    cmdClose.Enabled = True
    cmdChangePwd.Enabled = False
    txtPassword.Enabled = False
    Exit Sub
    End If
    '----------
    If cmdChangePwd.Caption = "Change User PWD/PRF" Then
    cmdChangePwd.Caption = "Save"
    cmdAddUser.Enabled = False
    cmdDeleteUser.Enabled = False
    cmdFindUser.Enabled = False
    cmdClose.Enabled = False
    cmdChangePwd.Enabled = True
    txtPassword.Enabled = True
    cboLevel.Enabled = True
    txtPassword.SetFocus
    Exit Sub
    End If
    '----------
    If cmdChangePwd.Caption = "Save" Then
    cmdChangePwd.Caption = "Change User PWD/PRF"
    cmdAddUser.Enabled = True
    cmdDeleteUser.Enabled = False
    cmdChangePwd.Enabled = False
    cmdFindUser.Enabled = True
    cmdClose.Enabled = True
    txtPassword.Enabled = False
    cboLevel.Enabled = False
    End If
    '----------
    dbTable.Edit
    '**********
    If txtUserName.Text = "ADMIN" And txtPassword.Text = "admin" Then
    dbTable!UserName = txtUserName.Text
    dbTable!Password = txtPassword.Text
    dbTable!Level = cboLevel.Text
    dbTable.Update
    cmdChangePwd.Enabled = False
    cmdClose.Enabled = True
    txtPassword.Enabled = False
    Exit Sub
    Else
    txtPassword.Text = KLEDll.Encrypt(txtPassword.Text)
    End If
    '**********
    dbTable!UserName = txtUserName.Text
    dbTable!Password = txtPassword.Text
    dbTable!Level = cboLevel.Text
    dbTable.Update
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDeleteUser_Click()
    cmdDeleteUser.Enabled = False
    cmdChangePwd.Enabled = False
    dbTable.Delete
    If dbTable.EOF Then dbTable.MoveFirst
    dbTable.MoveNext
End Sub

Private Sub cmdFindUser_Click()
    If dbTable.EOF Then dbTable.MoveFirst
    dbTable.MoveNext
    If dbTable.EOF Then dbTable.MoveFirst
    txtUserName.Text = dbTable!UserName
    txtPassword.Text = dbTable!Password
    cboLevel.Text = dbTable!Level
    '----------
    cmdDeleteUser.Enabled = True
    cmdChangePwd.Enabled = True
    cmdAddUser.Caption = "Add User"
    txtUserName.Enabled = False
    txtPassword.Enabled = False
    cboLevel.Enabled = False
End Sub

Private Sub cmdFindUser1_Click()
    If dbTable.EOF Then dbTable.MoveFirst
    dbTable.MoveNext
    If dbTable.EOF Then dbTable.MoveFirst
    txtUserName.Text = dbTable!UserName
    txtPassword.Text = dbTable!Password
    cboLevel.Text = dbTable!Level
End Sub

Private Sub Form_Load()
    frmUserProfile.Height = 2000
    '----------
    cboLevel.AddItem "User"
    cboLevel.AddItem "Administrator"
    cboLevel.ListIndex = 0
    '----------
    Set dbWorkspace = DBEngine.Workspaces(0)
    Set dbDatabase = dbWorkspace.OpenDatabase(App.Path & "\PSC.MDB", True, False, ";pwd=" & "PSC")
    Set dbTable = dbDatabase.OpenRecordset("Users", dbOpenTable)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Visible = True
    '---------
    If frmTemp!txtLevel.Text = "Administrator" Then
    frmMain!cmdUserProfile.Enabled = True
    frmMain!cmdChangePwd.Enabled = True
    End If
    '----------
    If frmTemp!txtLevel.Text = "User" Then
    frmMain!cmdUserProfile.Enabled = False
    frmMain!cmdChangePwd.Enabled = True
    End If
End Sub

Private Sub tmrFindUser_Timer()
    cmdFindUser1_Click
    If txtUserName.Text = frmTemp!txtUserName.Text Then
    tmrFindUser.Enabled = False
    End If
End Sub

Private Sub txtPassword_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(LCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtUsername_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtUsername_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

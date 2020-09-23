VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3630
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboLevel 
      Height          =   315
      Left            =   960
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   960
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1380
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   960
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1380
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox txtFileExist 
      Enabled         =   0   'False
      Height          =   285
      Left            =   840
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblLevel 
      AutoSize        =   -1  'True
      Caption         =   "Level:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   435
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      Caption         =   "Username:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   765
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Add Microsoft DAO 3.6 Object Library (Project -> References)
'Next Version. Encrypt Password.
'Add KLE.Dll to Project (Project -> References)
'-------------------------------------------------------------
Option Explicit
Private dbWorkspace As Workspace
Private dbDatabase As Database
Private dbTable As Recordset
Dim dbTableDef As TableDef
Dim dbUserName As Field
Dim dbPassword As Field
Dim dbLevel As Field
Public KLEDll As New prjEncPwd.clsEncPwd 'Info from the dll

Public Function mdbFileExists(strFile As String)
    mdbFileExists = False
    On Error Resume Next
    mdbFileExists = (FileLen(strFile) = FileLen(strFile))
End Function

Private Sub cboLevel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
    cmdOK_Click
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim frmLoginDB As Database
    Dim frmLoginRecordSet As Recordset
    
    Set frmLoginDB = OpenDatabase(App.Path & "\PSC.MDB", True, False, ";pwd=" & "PSC")
    Set frmLoginRecordSet = frmLoginDB.OpenRecordset("Users")
    
    '***
    If txtUserName.Text = "ADMIN" And txtPassword.Text = "admin" Then
    Else
    txtPassword.Text = KLEDll.Encrypt(txtPassword.Text)
    End If
    '***
    
    Do While Not frmLoginRecordSet.EOF
    If frmLoginRecordSet.Fields("UserName") = (txtUserName.Text) And _
    frmLoginRecordSet.Fields("Password") = (txtPassword.Text) And _
    frmLoginRecordSet.Fields("Level") = (cboLevel.Text) Then
    '----------
    frmTemp.Show
    frmTemp!txtUserName.Text = txtUserName.Text
    frmTemp!txtPassword.Text = txtPassword.Text
    frmTemp!txtLevel.Text = cboLevel.Text
    frmTemp.Visible = False
    '----------
    frmMain.Show
    If cboLevel.Text = "Administrator" Then
    frmMain!cmdUserProfile.Enabled = True
    frmMain!cmdChangePwd.Enabled = True
    End If
    '----------
    If cboLevel.Text = "User" Then
    frmMain!cmdUserProfile.Enabled = False
    frmMain!cmdChangePwd.Enabled = True
    End If
    '----------
    Unload Me
    Exit Sub
    Else
    frmLoginRecordSet.MoveNext
    End If
    Loop
    txtPassword.Text = ""
    MsgBox "Invalid Username or Password. Please Try Again.", vbOKOnly + vbCritical + vbSystemModal, "Phone Book - [Login]"
End Sub

Private Sub Form_Load()
    Me.Height = 2210
    '----------
    cboLevel.AddItem "Administrator"
    cboLevel.AddItem "User"
    cboLevel.ListIndex = 1
    '----------
    txtFileExist.Text = mdbFileExists(App.Path & "\PSC.MDB")
    If txtFileExist.Text = "True" Then
    DBEngine.CompactDatabase App.Path & "\PSC.MDB", App.Path & "\PSC1.MDB", dbLangGeneral, , ";pwd=" & "PSC"
    Kill App.Path & "\PSC.MDB"
    FileCopy App.Path & "\PSC1.MDB", App.Path & "\PSC.MDB"
    Kill App.Path & "\PSC1.MDB"
    Exit Sub
    End If
    '----------
    Set dbWorkspace = DBEngine.Workspaces(0)
    Set dbDatabase = dbWorkspace.CreateDatabase(App.Path & "\PSC.MDB", dbLangGeneral)
    Set dbTableDef = dbDatabase.CreateTableDef("Users")
    Set dbUserName = dbTableDef.CreateField("Username", dbText, 30)
    Set dbPassword = dbTableDef.CreateField("Password", dbText, 30)
    Set dbLevel = dbTableDef.CreateField("Level", dbText, 30)
    
    dbDatabase.NewPassword "", "PSC"
    
    dbTableDef.Fields.Append dbUserName
    dbTableDef.Fields.Append dbPassword
    dbTableDef.Fields.Append dbLevel
    dbDatabase.TableDefs.Append dbTableDef
    dbDatabase.Close
    '----------
    Set dbWorkspace = DBEngine.Workspaces(0)
    Set dbDatabase = dbWorkspace.OpenDatabase(App.Path & "\PSC.MDB", True, False, ";pwd=" & "PSC")
    Set dbTable = dbDatabase.OpenRecordset("Users", dbOpenTable)
    If dbTable.BOF And dbTable.EOF Then
    dbTable.AddNew
    dbTable!UserName = "ADMIN"
    dbTable!Password = "admin"
    dbTable!Level = "Administrator"
    dbTable.Update
    dbTable.MoveLast
    End If
    '----------
    txtFileExist.Text = mdbFileExists(App.Path & "\PSC.MDB")
End Sub

Private Sub txtPassword_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
    cmdOK_Click
    End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(LCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtUsername_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
    txtPassword.SetFocus
    End If
End Sub

Private Sub txtUsername_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

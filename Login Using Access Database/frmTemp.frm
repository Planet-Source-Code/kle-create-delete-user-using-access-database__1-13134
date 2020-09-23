VERSION 5.00
Begin VB.Form frmTemp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Temp"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3645
   ControlBox      =   0   'False
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   3645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtLevel 
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblLevel 
      AutoSize        =   -1  'True
      Caption         =   "Level:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   435
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      Caption         =   "Username:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   765
   End
End
Attribute VB_Name = "frmTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

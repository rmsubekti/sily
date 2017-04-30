VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2895
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710.462
   ScaleMode       =   0  'User
   ScaleWidth      =   3985.825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1530
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1845
      Width           =   2325
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   390
      Left            =   2340
      TabIndex        =   2
      Top             =   2340
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Login"
      Default         =   -1  'True
      Height          =   390
      Left            =   735
      TabIndex        =   1
      Top             =   2340
      Width           =   1140
   End
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1530
      TabIndex        =   0
      Top             =   1455
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   345
      TabIndex        =   7
      Top             =   1860
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   345
      TabIndex        =   6
      Top             =   1470
      Width           =   1080
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   900
      Left            =   120
      Picture         =   "frmLogin.frx":0000
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lang"
      BeginProperty Font 
         Name            =   "Script"
         Size            =   48
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   960
      TabIndex        =   5
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Laundry"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private password As String
Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdOK_Click()
    userlogin = txtUserName.Text
    password = txtPassword.Text
    rs.Open "select * from ulogin where uid='" & _
        userlogin & "' and password='" & password & "'", conn, adOpenForwardOnly, adLockReadOnly
        
    If rs.EOF Then
        MsgBox "Tidak ada pengguna dengan username " & userlogin & vbCrLf & _
                "atau password yang dimasukkan salah.", , "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
        
        rs.Close
        Set rs = Nothing
        
    Else
    
        akses = rs("akses")
        rs.Close
        Me.Hide
        Set rs = Nothing
        rs.Open "select * from karyawan where uid='" & userlogin & "'", conn
        
        UserName = rs("nama")
        MsgBox "Selamat datang " & UserName & "." & _
                vbCrLf & "Anda masuk sebagai " & akses & ".", _
                vbInformation, "Login Success"
        
        rs.Close
        Set rs = Nothing
        If UCase(akses) = "ADMIN" Then
            MDIAdmin.Show
        Else
            FrmKasir.Show
            MDIAdmin.Caption = "Dashboard Kasir"
        End If
        Me.Hide
    End If
    
End Sub

Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtUserName.SetFocus
End Sub



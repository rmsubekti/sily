VERSION 5.00
Begin VB.Form FrmAkun 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informasi Akun"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5505
   ScaleWidth      =   4845
   Begin VB.Frame Frame1 
      Caption         =   "Info Kontak"
      Height          =   2655
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   4575
      Begin VB.TextBox txtAlamat 
         Height          =   285
         Left            =   1320
         TabIndex        =   18
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txtNama 
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox txtNo 
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtJabatan 
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   1440
         Width           =   3015
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Simpan Kontak"
         Height          =   375
         Left            =   3000
         TabIndex        =   9
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Alamat"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Nama"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label label 
         Caption         =   "Jabatan"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "No. Telp"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ganti Password"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   4575
      Begin VB.TextBox txtLama 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtBaru 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtKonfirm 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CommandButton cmdUbah 
         Caption         =   "Ubah Password"
         Height          =   375
         Left            =   3000
         TabIndex        =   1
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Password Lama"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Password Baru"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Konfirmasi Password"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
      End
   End
   Begin VB.Label lblUsername 
      Alignment       =   2  'Center
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   14.25
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   16
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "FrmAkun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()
    rs.Open "select * from karyawan where nik='" & _
            userLogin & "'", conn, adOpenKeyset, adLockOptimistic
        rs!nama = txtNama.Text
        rs!telp = txtNo.Text
        rs!alamat = txtAlamat.Text
        rs!jabatan = txtJabatan.Text
    rs.Update
    rs.Close
    MsgBox "Kontak tersimpan", vbInformation, "Sukses"
End Sub

Private Sub cmdUbah_Click()
    If txtBaru.Text = txtKonfirm.Text Then
        rs.Open "select * from login_karyawan where username='" & _
        userLogin & "' and password='" & txtLama.Text & "'", conn, adOpenForwardOnly, adLockReadOnly
        
        If rs.EOF Then
            MsgBox "Password yang anda masukkan tidak sesuai dengan Akun " & userLogin, , "Password Salah"
            txtLama.SetFocus
            rs.Close
            Set rs = Nothing
        Else
            
            rs.Close
            Set rs = Nothing
            rs.Open "select * from login_karyawan where username='" & _
                userLogin & "'", conn, adOpenKeyset, adLockOptimistic
                rs!password = txtBaru.Text
            rs.Update
            rs.Close
            
            MsgBox "Password telah diganti untuk login selanjutnya", vbInformation, "Password Diganti"
        
        End If
    Else
        MsgBox "Silakan masukkan kembali password baru anda " & vbCrLf & _
            "untuk mengonfirmasi perubahan password.", vbCritical, "Konfirmasi password tidak cocok!"
    End If
End Sub

Private Sub Form_Load()
    lblUsername = userLogin
    rs.Open "select * from karyawan where nik = '" & userLogin & "'", conn, adOpenForwardOnly, adLockReadOnly
        If Not rs.EOF Then
            txtNama.Text = rs("nama")
            txtNo.Text = rs("telp")
            txtAlamat.Text = rs("alamat")
            txtJabatan.Text = rs("jabatan")
        End If
    rs.Close
End Sub



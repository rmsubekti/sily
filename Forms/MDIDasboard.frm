VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIDasboard 
   BackColor       =   &H8000000C&
   Caption         =   "Dasboard"
   ClientHeight    =   5100
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9225
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   2  'Align Bottom
      Height          =   630
      Left            =   0
      TabIndex        =   1
      Top             =   4230
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   4860
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnMaster 
      Caption         =   "&Master"
      Begin VB.Menu mnPaket 
         Caption         =   "Paket"
      End
      Begin VB.Menu mnPelanggan 
         Caption         =   "&Pelanggan"
      End
      Begin VB.Menu mnKaryawan 
         Caption         =   "&Karyawan"
      End
   End
   Begin VB.Menu mnTransaksi 
      Caption         =   "&Transaksi"
   End
   Begin VB.Menu mnLaporan 
      Caption         =   "&Laporan"
   End
   Begin VB.Menu mnSetting 
      Caption         =   "&Setting"
      Begin VB.Menu mnAkun 
         Caption         =   "Akun Profile"
      End
   End
End
Attribute VB_Name = "MDIDasboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnKaryawan_Click()
    FrmKaryawan.Show
End Sub

Private Sub mnPaket_Click()
    FrmPaket.Show
End Sub

Private Sub mnPelanggan_Click()
    FrmPelanggan.Show
End Sub

Private Sub mnTransaksi_Click()
    FrmTransaksi.Show
End Sub

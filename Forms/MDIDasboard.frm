VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIDasboard 
   BackColor       =   &H8000000C&
   Caption         =   "Dasboard Admin"
   ClientHeight    =   5100
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9225
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2520
      Top             =   4800
   End
   Begin MSComctlLib.StatusBar sbInfo 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4845
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
Private Sub MDIForm_Load()
    Dim p As Panel
    sbInfo.Panels(1).Text = userName
    sbInfo.Panels(1).ToolTipText = "Masuk sebagai " & userName
    Set p = sbInfo.Panels.Add(2, , " :: Selamat Datang Di LANG Laundry :: Terima Kasih Atas Kepercayaan Anda Berlangganan Jasa Kami :: Kami Selalu Siap Melayani dan Memastikan Anda Mendapatkan Pelayanan Terbaik Dari Kami ")
    Set p = sbInfo.Panels.Add(3, , , sbrDate)
    Set p = sbInfo.Panels.Add(4, , , sbrTime)
    sbInfo.Panels(2).Style = sbrText
    sbInfo.Panels(2).MinWidth = 10000
    sbInfo.Panels(1).MinWidth = 2201
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    frmLogin.Show
End Sub

Private Sub mnAkun_Click()
    FrmAkun.Show
End Sub

Private Sub mnKaryawan_Click()
    FrmKaryawan.Show
End Sub

Private Sub mnLaporan_Click()
    FrmLaporan.Show
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
Private Sub Timer1_Timer()
sbInfo.Panels(2).Text = Right(sbInfo.Panels(2).Text, Len(sbInfo.Panels(2).Text) - 1) + Left(sbInfo.Panels(2).Text, 1)
End Sub


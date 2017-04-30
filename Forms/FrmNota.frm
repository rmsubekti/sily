VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form FrmNota 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Nota"
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1935
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   6375
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   900
         Left            =   120
         Picture         =   "FrmNota.frx":0000
         Top             =   120
         Width           =   900
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Lang"
         BeginProperty Font 
            Name            =   "Script"
            Size            =   47.25
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1080
         TabIndex        =   22
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Laundry"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   21
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Tambakboyo  Rt  22  Rw  61  No  758  Condongcatur"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1200
         Width           =   5295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Depok Sleman Hp. 0877 3884 7881, 0852 9247 5400"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   4935
      End
      Begin VB.Line Line1 
         DrawMode        =   5  'Not Copy Pen
         X1              =   120
         X2              =   6360
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line2 
         DrawMode        =   5  'Not Copy Pen
         X1              =   120
         X2              =   6360
         Y1              =   1830
         Y2              =   1830
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1575
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   6375
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Terima  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Ambil  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   16
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama   :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   145
         TabIndex        =   15
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "No.Tlp  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1215
      End
      Begin VB.Line Line3 
         X1              =   960
         X2              =   6120
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Line Line4 
         X1              =   960
         X2              =   6120
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Label lblTerima 
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   13
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label lblAmbil 
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   12
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label lblNama 
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   11
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label lblNo 
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   840
         Width           =   5175
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblAlamat 
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   1200
         Width           =   5175
      End
      Begin VB.Line Line5 
         X1              =   960
         X2              =   6120
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Line Line9 
         X1              =   960
         X2              =   3000
         Y1              =   320
         Y2              =   320
      End
      Begin VB.Line Line10 
         X1              =   3960
         X2              =   6120
         Y1              =   320
         Y2              =   320
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000009&
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   7680
      Width           =   6375
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "1. Garansi cucian yang kurang bersih/rapi akan diproses ulang"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   6255
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "2. Kami tidak bertanggung jawab susut/luntur karena sifat bahannya"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   6255
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "3. Barang hilang/rusak akan diganti maks 10x tarif Laundry"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   6255
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "4. Klaim hanya berlaku 1x24 jam setelah barang diambil dengan menyertakan nota asli"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   6255
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "5. Barang yang tidak diambil 1 bulan diluar tanggung jawab Lang Laundry"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   6255
      End
   End
   Begin MSComctlLib.ListView lstOrder 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PERHATIAN !!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   28
      Top             =   7440
      Width           =   4455
   End
   Begin VB.Line Line6 
      X1              =   4080
      X2              =   6360
      Y1              =   6810
      Y2              =   6810
   End
   Begin VB.Line Line7 
      X1              =   4080
      X2              =   6360
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Rp."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   27
      Top             =   6480
      Width           =   615
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "Label18"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   26
      Top             =   6480
      Width           =   2295
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Terimakasih atas kepercayaan anda"
      Height          =   255
      Left            =   1080
      TabIndex        =   25
      Top             =   7080
      Width           =   4695
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Nota :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblNota 
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   23
      Top             =   2160
      Width           =   5175
   End
   Begin VB.Line Line8 
      X1              =   1200
      X2              =   6240
      Y1              =   2340
      Y2              =   2340
   End
End
Attribute VB_Name = "FrmNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    lstOrder.ColumnHeaders.Add , , "Paket", 2400
    lstOrder.ColumnHeaders.Add , , "Jumlah", 1400
    lstOrder.ColumnHeaders.Add , , "Tarif", 2500
    
    lblNota.Caption = id_transaksi
    rs.Open "select * from pelanggan where id_pelanggan = '" & id_pelanggan & "'", conn, adOpenForwardOnly, adLockReadOnly
    Do Until rs.EOF
        lblNama.Caption = rs("nama")
        lblNo.Caption = rs("telp")
        lblAlamat.Caption = rs("alamat")
        rs.MoveNext
    Loop
    rs.Close
    
    rs.Open "select * from transaksi where id_transaksi = '" & id_transaksi & "'", conn, adOpenForwardOnly, adLockReadOnly
    Do Until rs.EOF
        lblTerima.Caption = Format(rs("tgl_terima"), "DD - MM - YYYY")
        lblAmbil.Caption = Format(rs("tgl_ambil"), "DD - MM - YYYY")
        lblTotal.Caption = rs("biaya")
        rs.MoveNext
    Loop
    rs.Close
    
    Dim lsItem As ListItem
    rs.Open "select * from det_transaksi inner join paket ON det_transaksi.id_paket=paket.id_paket where det_transaksi.id_transaksi = '" & id_transaksi & "'", conn, adOpenForwardOnly, adLockReadOnly
    Do Until rs.EOF
        Set lsItem = lstOrder.ListItems.Add(, , rs("nama"))
            lsItem.SubItems(1) = rs("jumlah")
            lsItem.SubItems(2) = rs("total")
        rs.MoveNext
    Loop
    rs.Close
    
End Sub


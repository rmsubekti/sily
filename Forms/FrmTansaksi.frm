VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form FrmTransaksi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaksi"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   11085
   Begin VB.CommandButton cmdUlang 
      Caption         =   "Print Ulang"
      Height          =   375
      Left            =   8400
      TabIndex        =   21
      Top             =   4080
      Width           =   2535
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "Batal"
      Height          =   330
      Left            =   9720
      TabIndex        =   17
      Top             =   1800
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   330
      Left            =   8400
      TabIndex        =   16
      Top             =   1800
      Width           =   1200
   End
   Begin MSComctlLib.ListView lstPelanggan 
      Height          =   1455
      Left            =   8280
      TabIndex        =   15
      Top             =   240
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   2566
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lstTransaksi 
      Height          =   2175
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   3836
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data Pelanggan"
      Height          =   2055
      Left            =   5160
      TabIndex        =   6
      Top             =   120
      Width           =   3015
      Begin VB.TextBox txtAlamat 
         Height          =   300
         Left            =   840
         TabIndex        =   13
         Top             =   1440
         Width           =   2000
      End
      Begin VB.TextBox txtNo 
         Height          =   300
         Left            =   840
         TabIndex        =   12
         Top             =   1080
         Width           =   2000
      End
      Begin VB.TextBox txtNama 
         Height          =   300
         Left            =   840
         TabIndex        =   11
         Top             =   720
         Width           =   2000
      End
      Begin VB.Label lblKode 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   840
         TabIndex        =   22
         Top             =   360
         Width           =   2000
      End
      Begin VB.Label Label5 
         Caption         =   "Kode"
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Alamat"
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "No.Telp"
         Height          =   375
         Left            =   180
         TabIndex        =   8
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Nama"
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pilih Paket Laundry"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtJumlah 
         Height          =   300
         Left            =   840
         TabIndex        =   23
         Top             =   1590
         Width           =   1575
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   330
         Left            =   3720
         TabIndex        =   5
         Top             =   1590
         Width           =   1100
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   330
         Left            =   2520
         TabIndex        =   4
         Top             =   1590
         Width           =   1100
      End
      Begin MSComctlLib.ListView lstPaket 
         Height          =   1215
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   2143
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lstPilih 
         Height          =   1215
         Left            =   2520
         TabIndex        =   2
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   2143
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label1 
         Caption         =   "Jumlah"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1600
         Width           =   615
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Rp."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   20
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      TabIndex        =   19
      Top             =   3180
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Total Bayar"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   18
      Top             =   2520
      Width           =   2535
   End
End
Attribute VB_Name = "FrmTransaksi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

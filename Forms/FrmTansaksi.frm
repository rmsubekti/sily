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
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      Caption         =   "0"
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
Option Explicit
Dim jumlah, total, tarif As Integer
Dim lsItem As ListItem

Private Sub cmdAdd_Click()
    If txtJumlah.Text = "" Then
        MsgBox "Silakan masukkan jumlah paket yang di pesan", vbInformation, "Jumlah Belum diisi"
        txtJumlah.SetFocus
    Else
        tarif = Val(lstPaket.SelectedItem.SubItems(3)) * Val(txtJumlah.Text)
        
        Set lsItem = lstPilih.ListItems.Add(, , lstPaket.SelectedItem.Text)
        lsItem.SubItems(1) = lstPaket.SelectedItem.SubItems(1)
        lsItem.SubItems(2) = lstPaket.SelectedItem.SubItems(2)
        lsItem.SubItems(3) = tarif
        lsItem.SubItems(4) = txtJumlah.Text
        
        total = total + tarif
        lblTotal.Caption = total
        lstPaket.ListItems.Remove lstPaket.SelectedItem.Index
    End If
End Sub

Private Sub cmdBatal_Click()
    lstPilih.ListItems.Clear
    lblTotal.Caption = 0
    showPaket
    lblKode.Caption = ""
    txtNama.Text = ""
    txtNo.Text = ""
    txtAlamat.Text = ""
End Sub

Private Sub cmdPrint_Click()
    saveTransaksi
    FrmNota.Show
End Sub

Private Sub cmdRemove_Click()
    tarif = Val(lstPilih.SelectedItem.SubItems(3)) / Val(lstPilih.SelectedItem.SubItems(4))
    
    Set lsItem = lstPaket.ListItems.Add(, , lstPilih.SelectedItem.Text)
    lsItem.SubItems(1) = lstPilih.SelectedItem.SubItems(1)
    lsItem.SubItems(2) = lstPilih.SelectedItem.SubItems(2)
    lsItem.SubItems(3) = tarif
    
    total = total - Val(lstPilih.SelectedItem.SubItems(3))
    lblTotal.Caption = total
    lstPilih.ListItems.Remove lstPilih.SelectedItem.Index
End Sub

Private Sub cmdUlang_Click()
    id_transaksi = lstTransaksi.SelectedItem.Text
    id_pelanggan = lstTransaksi.SelectedItem.SubItems(5)
    FrmNota.Show
End Sub

Private Sub Form_Load()
    lstPaket.ColumnHeaders.Add , , "ID", 0
    lstPaket.ColumnHeaders.Add , , "Paket", 2120
    lstPaket.ColumnHeaders.Add , , "Satuan", 0
    lstPaket.ColumnHeaders.Add , , "Tarif", 0
    
    lstPilih.ColumnHeaders.Add , , "ID", 0
    lstPilih.ColumnHeaders.Add , , "Paket Diambil", 2120
    lstPilih.ColumnHeaders.Add , , "Satuan", 0
    lstPilih.ColumnHeaders.Add , , "Tarif", 0
    lstPilih.ColumnHeaders.Add , , "Jumlah", 0
    
    lstPelanggan.ColumnHeaders.Add , , "ID", 0
    lstPelanggan.ColumnHeaders.Add , , "Nama Pelanggan", 2500
    lstPelanggan.ColumnHeaders.Add , , "No Telp", 0
    lstPelanggan.ColumnHeaders.Add , , "Alamat", 0
    
    lstTransaksi.ColumnHeaders.Add , , "ID", 850
    lstTransaksi.ColumnHeaders.Add , , "Nama Pelanggan", 1900
    lstTransaksi.ColumnHeaders.Add , , "Total Bayar", 1500
    lstTransaksi.ColumnHeaders.Add , , "Tanggal Diterima", 1800
    lstTransaksi.ColumnHeaders.Add , , "Tanggal Diambil", 1800
    lstTransaksi.ColumnHeaders.Add , , "id_pelanggan", 0
    Call getKoneksi
    showPaket
    showTransaksi
End Sub
Private Sub showPaket()
    lstPaket.ListItems.Clear
    Dim lsItem As ListItem
    rs.Open "paket", conn, adOpenForwardOnly, adLockReadOnly
    If rs.EOF Then
        lstPaket.ListItems.Clear
    Else
        Do Until rs.EOF
            Set lsItem = lstPaket.ListItems.Add(, , rs("id_paket"))
                lsItem.SubItems(1) = rs("nama")
                lsItem.SubItems(2) = rs("satuan")
                lsItem.SubItems(3) = rs("tarif")
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub showTransaksi()
    lstTransaksi.ListItems.Clear
    If rs.State = adStateOpen Then
        rs.Close
    End If
    rs.Open "select * from transaksi inner join pelanggan ON transaksi.id_pelanggan=pelanggan.id_pelanggan; ", conn, adOpenForwardOnly, adLockReadOnly
    Do Until rs.EOF
        Set lsItem = lstTransaksi.ListItems.Add(, , rs("id_transaksi"))
            lsItem.SubItems(1) = rs("nama")
            lsItem.SubItems(2) = rs("biaya")
            lsItem.SubItems(3) = rs("tgl_terima")
            lsItem.SubItems(4) = rs("tgl_ambil")
            lsItem.SubItems(5) = rs("id_pelanggan")
        rs.MoveNext
    Loop
    rs.Close
End Sub
Private Sub searchPelanggan()
    lstPelanggan.ListItems.Clear
    If txtNama.Text <> "" Then
        Dim lsItem As ListItem
        rs.Open "select * from pelanggan where nama like '%" & txtNama.Text & "%'", _
                    conn, adOpenForwardOnly, adLockReadOnly
        If rs.EOF Then
        
        Else
            Do Until rs.EOF
                Set lsItem = lstPelanggan.ListItems.Add(, , rs("id_pelanggan"))
                    lsItem.SubItems(1) = rs("nama")
                    lsItem.SubItems(2) = rs("telp")
                    lsItem.SubItems(3) = rs("alamat")
                rs.MoveNext
            Loop
        End If
        rs.Close
        Set rs = Nothing
    End If
End Sub
Private Sub lstPelanggan_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lblKode.Caption = lstPelanggan.SelectedItem.Text
    txtNama.Text = lstPelanggan.SelectedItem.SubItems(1)
    txtNo.Text = lstPelanggan.SelectedItem.SubItems(2)
    txtAlamat.Text = lstPelanggan.SelectedItem.SubItems(3)
End Sub
Private Sub txtNama_Change()
    searchPelanggan
End Sub

Private Sub saveTransaksi()
    id_transaksi = generateIDTransaksi
    If lblKode.Caption <> "" Then
        id_pelanggan = lblKode.Caption
        conn.Execute "insert into transaksi values('" & id_transaksi & "','" & id_pelanggan & "','" & total & _
            "',cast(getdate() as datetime), cast(getdate() + 3 as datetime))"
        
        For Each lsItem In lstPilih.ListItems
            conn.Execute "insert into det_transaksi values('" & id_transaksi & "','" & _
                lsItem.Text & "','" & lsItem.SubItems(4) & "','" & lsItem.SubItems(3) & "')"
        Next
    Else
        id_pelanggan = generateIDPelanggan
        conn.Execute "insert into pelanggan values('" & id_pelanggan & "','" & _
            txtNama.Text & "','" & txtNo.Text & "','" & txtAlamat.Text & "')"
            
        conn.Execute "insert into transaksi values('" & id_transaksi & "','" & id_pelanggan & "','" & total & _
            "',cast(getdate() as datetime), cast(getdate() + 3 as datetime))"
        
        For Each lsItem In lstPilih.ListItems
            conn.Execute "insert into det_transaksi values('" & id_transaksi & "','" & _
                lsItem.Text & "','" & lsItem.SubItems(4) & "','" & lsItem.SubItems(3) & "')"
        Next
    End If
    showTransaksi
End Sub



VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmPelanggan 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Pelanggan"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   10815
   Begin MSAdodcLib.Adodc adoPelanggan 
      Height          =   330
      Left            =   3840
      Top             =   3600
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Data Pelanggan"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdNew 
      Appearance      =   0  'Flat
      Caption         =   "New"
      Height          =   330
      Left            =   120
      TabIndex        =   15
      Top             =   3610
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      Caption         =   "Save"
      Height          =   330
      Left            =   1320
      TabIndex        =   14
      Top             =   3610
      Width           =   1100
   End
   Begin VB.CommandButton cmdDelete 
      Appearance      =   0  'Flat
      Caption         =   "Delete"
      Height          =   330
      Left            =   2520
      TabIndex        =   13
      Top             =   3610
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "List Data Pelanggan"
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   3840
      TabIndex        =   9
      Top             =   120
      Width           =   6855
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   5055
      End
      Begin VB.CommandButton cmdSearch 
         Appearance      =   0  'Flat
         Caption         =   "Search"
         Height          =   300
         Left            =   5280
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid dataPelanggan 
         Bindings        =   "FrmPelanggan.frx":0000
         Height          =   2175
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   3836
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         Appearance      =   0
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Input Data Pelanggan"
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.TextBox txtNama 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtAlamat 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Kode"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Nama"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Alamat"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "No. Telp"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label txtKode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   360
         Width           =   2175
      End
   End
End
Attribute VB_Name = "FrmPelanggan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()
    If MsgBox("Anda akan menghapus " & txtNama.Text & _
        " dari data pelanggan.", vbOKCancel, "Hapus " & txtNama.Text) = vbOK Then
        sql = "delete pelanggan where id_pelanggan ='" & txtKode.Caption & "'"
        conn.Execute (sql)
        
        refreshDataGrid
        clearText
        needSave = False
    End If
End Sub

Private Sub cmdNew_Click()
    forgotSave
    needSave = True
    clearText
    incrementAngka
    cmdDelete.Enabled = False
End Sub

Private Sub cmdSave_Click()
    saveChanges
End Sub

Private Sub cmdSearch_Click()
    'Dim searchQuery As String
    'searchQuery = IIf(txtSearch.Text <> "", "'%" & txtSearch.Text & "%'", "'%'")
    adoPelanggan.RecordSource = "select * from pelanggan where nama like '%" & txtSearch.Text & "%';"
    
    adoPelanggan.Refresh
    If adoPelanggan.Recordset.BOF Then
        MsgBox ("Pelanggan dengan nama " & txtSearch.Text & " tidak ada.")
        refreshDataGrid
    End If
End Sub

Private Sub dataPelanggan_Click()
    forgotSave
    With adoPelanggan
        txtKode.Caption = .Recordset(0)
        txtNama.Text = .Recordset(1)
        txtNo.Text = .Recordset(2)
        txtAlamat.Text = .Recordset(3)
    End With
    needSave = False
End Sub

Private Sub Form_Load()
    Call getKoneksi
    refreshDataGrid
    clearText
    needSave = False
    dataPelanggan.Columns(0).Width = 800
End Sub

Private Sub Form_Unload(Cancel As Integer)
    forgotSave
End Sub

Private Sub txtAlamat_Change()
    needSave = True
End Sub

Private Sub txtNama_Change()
    needSave = True
End Sub

Private Sub txtNo_Change()
    needSave = True
End Sub
Private Sub refreshDataGrid()
    sql = "pelanggan"
    adoPelanggan.ConnectionString = konek
    adoPelanggan.RecordSource = sql
    adoPelanggan.Refresh
    Set dataPelanggan.DataSource = adoPelanggan
End Sub
Private Sub clearText()
    On Error Resume Next
    txtKode.Caption = ""
    txtNama.Text = ""
    txtAlamat.Text = ""
    txtNo.Text = ""
    
    txtNama.SetFocus
End Sub
Private Sub incrementAngka()
    Dim a As Integer
    sql = "select max(right(id_pelanggan,6)) from pelanggan"
    Set rs = conn.Execute(sql)
    a = IIf(rs(0) <> "NULL", rs(0) + 1, 1)
    If Val(a) < 10 Then
        txtKode.Caption = "P00000" & a
    ElseIf Val(a) > 10 And Val(a) < 100 Then
        txtKode.Caption = "P0000" & a
    ElseIf Val(a) > 100 And Val(a) < 1000 Then
        txtKode.Caption = "P000" & a
    ElseIf Val(a) > 1000 And Val(a) < 10000 Then
        txtKode.Caption = "P00" & a
    ElseIf Val(a) > 10000 And Val(a) < 100000 Then
        txtKode.Caption = "P0" & a
    Else
        txtKode.Caption = "P" & a
    End If
End Sub

Private Sub forgotSave()
    If needSave Then
        If MsgBox("Data yang diubah belum tersimpan. " & vbCrLf & _
            "Simpan sekarang ?", vbYesNo, "Konfirmasi") = vbYes Then
            saveChanges
        End If
        needSave = False
    End If
End Sub

Function isTextEmpty() As Boolean
    If txtNama.Text = "" Or txtNo.Text = "" Or txtAlamat.Text = "" Then isTextEmpty = True
End Function

Private Sub saveChanges()
    If isTextEmpty Then GoTo textKosong
    sql = "select * from pelanggan where id_pelanggan='" & txtKode.Caption & "'"
    Set rs = conn.Execute(sql)
    If rs.EOF Then
        sql = "insert into pelanggan values('" & _
            txtKode.Caption & "','" & _
            txtNama.Text & "','" & _
            txtNo.Text & "','" & _
            txtAlamat.Text & "')"
        Set rs = conn.Execute(sql)
    Else
        sql = "update pelanggan set id_pelanggan ='" & txtKode.Caption & _
            "', nama='" & txtNama.Text & _
            "', telp='" & txtNo.Text & _
            "', alamat='" & txtAlamat.Text & "'"
        Set rs = conn.Execute(sql)
    End If
    refreshDataGrid
    clearText
    needSave = False
Exit Sub
textKosong:
    MsgBox "Silakan masukkan informasi lengkap pelanggan.", vbCritical, "Input Kosong"
End Sub


Private Sub txtSearch_Change()
    adoPelanggan.RecordSource = "select * from pelanggan where nama like '%" & txtSearch.Text & "%';"
    
    adoPelanggan.Refresh
    If adoPelanggan.Recordset.BOF Then
        MsgBox ("Pelanggan dengan nama " & txtSearch.Text & " tidak ada.")
        refreshDataGrid
    End If
End Sub

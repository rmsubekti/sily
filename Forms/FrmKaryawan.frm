VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmKaryawan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Karyawan"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   10650
   Begin MSAdodcLib.Adodc dcKaryawan 
      Height          =   330
      Left            =   3840
      Top             =   3600
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"FrmKaryawan.frx":0000
      OLEDBString     =   $"FrmKaryawan.frx":008C
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "karyawan"
      Caption         =   "Data Karyawan"
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
   Begin VB.Frame Frame2 
      Caption         =   "List Data Karyawan"
      Height          =   3255
      Left            =   3840
      TabIndex        =   18
      Top             =   120
      Width           =   6615
      Begin MSDataGridLib.DataGrid dgKaryawan 
         Bindings        =   "FrmKaryawan.frx":0118
         Height          =   2175
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3836
         _Version        =   393216
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
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   320
         Left            =   5400
         TabIndex        =   20
         Top             =   340
         Width           =   975
      End
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   240
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2520
      TabIndex        =   16
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   1320
      TabIndex        =   15
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input Data Karyawan"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.TextBox txtJabatan 
         Height          =   285
         Left            =   1080
         TabIndex        =   13
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox txtAkses 
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Top             =   2520
         Width           =   2175
      End
      Begin VB.TextBox txtNo 
         Height          =   285
         Left            =   1080
         TabIndex        =   10
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtAlamat 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtNama 
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblNik 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Akses"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Password"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Jabatan"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "No. Telp"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Alamat"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Nama"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "NIK"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmKaryawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

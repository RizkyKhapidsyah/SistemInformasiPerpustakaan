VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.ocx"
Begin VB.Form FormPengembalianBuku 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pengembalian Buku"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10920
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormPengembalianBuku.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   10920
   Begin VB.CommandButton cmRefresh 
      Caption         =   "&Refresh"
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton cmHapus 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton cmFilter 
      Caption         =   "&Filter"
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton cmSorot 
      Caption         =   "&Sorot"
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton cmCari 
      Caption         =   "&Cari"
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton cmEdit 
      Caption         =   "&Edit"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   8400
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1815
      Left            =   120
      TabIndex        =   62
      Top             =   6480
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   3201
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   20
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
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
            LCID            =   1057
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
            LCID            =   1057
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
   Begin MSAdodcLib.Adodc AdodcPeminjamanBuku 
      Height          =   330
      Left            =   10800
      Top             =   9120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Appearance      =   1
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmReset 
      Caption         =   "&Reset"
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton cmTutup 
      Caption         =   "&Tutup"
      Height          =   495
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton cmSimpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buku"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   6255
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   5895
      Begin VB.TextBox textTahunPinjam 
         Alignment       =   2  'Center
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   3660
         Width           =   1005
      End
      Begin VB.TextBox textBulanPinjam 
         Alignment       =   2  'Center
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   3660
         Width           =   765
      End
      Begin VB.TextBox textTanggalPinjam 
         Alignment       =   2  'Center
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   3660
         Width           =   765
      End
      Begin VB.TextBox textHariPinjam 
         Alignment       =   2  'Center
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   3660
         Width           =   1245
      End
      Begin VB.TextBox textWaktuPeminjaman 
         Alignment       =   2  'Center
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   5760
         Width           =   3435
      End
      Begin VB.TextBox textAdminSaatPinjam 
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   5340
         Width           =   3075
      End
      Begin VB.TextBox textPengarang 
         DataField       =   "Pengarang"
         DataSource      =   "adcBuku"
         Height          =   790
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   1500
         Width           =   4395
      End
      Begin VB.TextBox textKeterangan 
         DataField       =   "Keterangan"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Top             =   4920
         Width           =   4395
      End
      Begin VB.TextBox textJumlah 
         Alignment       =   1  'Right Justify
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   4080
         Width           =   2085
      End
      Begin VB.TextBox textCetakanKe 
         Alignment       =   2  'Center
         DataField       =   "Cetakan"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   3240
         Width           =   2085
      End
      Begin VB.TextBox textTahunTerbit 
         Alignment       =   2  'Center
         DataField       =   "TahunTerbit"
         DataSource      =   "adcBuku"
         Height          =   450
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   2760
         Width           =   2085
      End
      Begin VB.TextBox textPenerbit 
         DataField       =   "Penerbit"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   2340
         Width           =   4155
      End
      Begin VB.TextBox textKodeBuku 
         DataField       =   "Judul"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   660
         Width           =   2115
      End
      Begin VB.ComboBox cmbJudulBuku 
         Height          =   390
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   240
         Width           =   4335
      End
      Begin VB.TextBox textLamaPinjam 
         Alignment       =   2  'Center
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   4500
         Width           =   1245
      End
      Begin VB.TextBox textKategori 
         DataField       =   "Judul"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1080
         Width           =   2475
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Wkt Peminjaman"
         Height          =   270
         Left            =   120
         TabIndex        =   46
         Top             =   5760
         Width           =   1080
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Admin Saat Pinjam"
         Height          =   270
         Left            =   120
         TabIndex        =   44
         Top             =   5340
         Width           =   1185
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cetakan ke"
         Height          =   270
         Left            =   120
         TabIndex        =   41
         Top             =   3260
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Judul Buku"
         Height          =   270
         Left            =   120
         TabIndex        =   40
         Top             =   300
         Width           =   705
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Kategori"
         Height          =   270
         Left            =   120
         TabIndex        =   39
         Top             =   1140
         Width           =   540
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Pengarang"
         Height          =   270
         Left            =   120
         TabIndex        =   38
         Top             =   1560
         Width           =   690
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         Height          =   270
         Left            =   120
         TabIndex        =   37
         Top             =   4920
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah"
         Height          =   270
         Left            =   120
         TabIndex        =   36
         Top             =   4100
         Width           =   465
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Pinjam"
         Height          =   270
         Left            =   120
         TabIndex        =   35
         Top             =   3720
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tahun Terbit"
         Height          =   270
         Left            =   120
         TabIndex        =   34
         Top             =   2820
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Penerbit"
         Height          =   270
         Left            =   120
         TabIndex        =   33
         Top             =   2400
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Buku"
         Height          =   270
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Lama Pinjam"
         Height          =   270
         Left            =   120
         TabIndex        =   31
         Top             =   4500
         Width           =   810
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Keterangan Peminjam"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3255
      Left            =   6120
      TabIndex        =   13
      Top             =   120
      Width           =   4695
      Begin VB.TextBox textNIA 
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   480
         Width           =   2925
      End
      Begin VB.TextBox textStatusPendidikan 
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   2715
         Width           =   2925
      End
      Begin VB.TextBox textNamaPeminjam 
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   945
         Width           =   2925
      End
      Begin VB.TextBox textAlamatPeminjam 
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   870
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   1380
         Width           =   3165
      End
      Begin VB.TextBox textNoTelpPeminjam 
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   2280
         Width           =   2925
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "NIA"
         Height          =   270
         Left            =   120
         TabIndex        =   70
         Top             =   495
         Width           =   210
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Peminjam"
         Height          =   270
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1020
      End
      Begin VB.Label labelalamatpeminjam 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat Peminjam"
         Height          =   270
         Left            =   120
         TabIndex        =   18
         Top             =   1395
         Width           =   1095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No. Telp Peminjam"
         Height          =   270
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   1185
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Status Pendidikan"
         Height          =   270
         Left            =   120
         TabIndex        =   16
         Top             =   2715
         Width           =   1125
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Data Admin"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2535
      Left            =   6120
      TabIndex        =   0
      Top             =   3420
      Width           =   4695
      Begin VB.TextBox textTahunPengembalian 
         Alignment       =   2  'Center
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   1200
         Width           =   1605
      End
      Begin VB.TextBox textBulanPengembalian 
         Alignment       =   2  'Center
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   1200
         Width           =   645
      End
      Begin VB.TextBox textTanggalPengembalian 
         Alignment       =   2  'Center
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   1200
         Width           =   525
      End
      Begin VB.ComboBox cmbHari 
         Height          =   390
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2070
         Width           =   2175
      End
      Begin VB.TextBox textJam 
         Alignment       =   2  'Center
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1640
         Width           =   525
      End
      Begin VB.TextBox textNamaAdmin 
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         TabIndex        =   4
         Top             =   360
         Width           =   3075
      End
      Begin VB.TextBox textMenit 
         Alignment       =   2  'Center
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1640
         Width           =   645
      End
      Begin VB.TextBox textDetik 
         Alignment       =   2  'Center
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1640
         Width           =   645
      End
      Begin VB.TextBox textBagian 
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         TabIndex        =   1
         Top             =   780
         Width           =   1605
      End
      Begin VB.Timer TimerWaktuPengembalian 
         Interval        =   10
         Left            =   4200
         Top             =   1680
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2880
         TabIndex        =   52
         Top             =   1200
         Width           =   45
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2040
         TabIndex        =   51
         Top             =   1200
         Width           =   45
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Pengembalian"
         Height          =   270
         Left            =   120
         TabIndex        =   50
         Top             =   1200
         Width           =   1110
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Hari"
         Height          =   270
         Left            =   120
         TabIndex        =   12
         Top             =   2070
         Width           =   270
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Wkt Pengembalian"
         Height          =   270
         Left            =   120
         TabIndex        =   11
         Top             =   1640
         Width           =   1185
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Admin"
         Height          =   270
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2040
         TabIndex        =   9
         Top             =   1640
         Width           =   45
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2880
         TabIndex        =   8
         Top             =   1640
         Width           =   45
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Bagian"
         Height          =   270
         Left            =   120
         TabIndex        =   7
         Top             =   780
         Width           =   420
      End
   End
   Begin MSAdodcLib.Adodc AdodcUtama 
      Height          =   330
      Left            =   10200
      Top             =   9360
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Appearance      =   1
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   6120
      X2              =   7560
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Created by Dwi Pradana"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   270
      Left            =   7680
      TabIndex        =   42
      Top             =   6120
      Width           =   1620
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      X1              =   9480
      X2              =   10800
      Y1              =   6240
      Y2              =   6240
   End
End
Attribute VB_Name = "FormPengembalianBuku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
NyambunggUtama
    With AdodcUtama
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * From tbPengembalianBuku order by JudulBuku asc;"
        Set DataGrid1.DataSource = AdodcUtama
        .Refresh
    End With
    With AdodcPeminjamanBuku
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * from tbPeminjamanBuku order by JudulBuku asc;"
        .Refresh
    End With
    For Each Objek In Me
        If TypeName(Objek) = "TextBox" Then
            With Objek
                .Text = ""
                .MaxLength = 254
            End With
        End If
    Next
    'MEMASUKKAN DATABASE NAMA BUKU KE CMBJUDULBUKU
    AdodcPeminjamanBuku.Refresh
    cmbJudulBuku.Clear
    Do Until AdodcPeminjamanBuku.Recordset.EOF
        cmbJudulBuku.AddItem AdodcPeminjamanBuku.Recordset.Fields(5).Value, 0
        AdodcPeminjamanBuku.Recordset.MoveNext
    Loop
    AdodcPeminjamanBuku.Refresh
    cmbJudulBuku.ListIndex = 0
    With cmbHari
        .Clear
        .AddItem "Minggu", 0
        .AddItem "Senin", 1
        .AddItem "Selasa", 2
        .AddItem "Rabu", 3
        .AddItem "Kamis", 4
        .AddItem "Jumat", 5
        .AddItem "Sabtu", 6
        .ListIndex = 1
    End With
    With DataGrid1
        .Columns(0).Width = 2280.189
        .Columns(1).Width = 840.189
        .Columns(2).Width = 764.7874
        .Columns(3).Width = 1365.165
        .Columns(4).Width = 1094.74
        .Columns(5).Width = 975.1182
        .Columns(6).Width = 884.9764
        .Columns(7).Width = 870.2363
        .Columns(8).Width = 1110.047
        .Columns(9).Width = 945.0709
        .Columns(10).Width = 1049.953
        .Columns(11).Width = 585.0709
        .Columns(12).Width = 959.8111
        .Columns(13).Width = 840.189
        .Columns(14).Width = 1349.858
        .Columns(15).Width = 1425.26
        .Columns(16).Width = 1200.189
        .Columns(17).Width = 1260.284
        .Columns(18).Width = 1319.811
        .Columns(19).Width = 1365.165
        .Columns(20).Width = 1005.165
        .Columns(21).Width = 675.2126
        .Columns(22).Width = 1649.764
        .Columns(23).Width = 1454.74
        .Columns(24).Width = 1470.047
        .Columns(25).Width = 1349.858
        .Columns(26).Width = 1425.26
        .Columns(27).Width = 1454.74
        .Columns(28).Width = 629.8583
    End With
End Sub
Sub KosongkanInput()
    For Each Objek In Me
        If TypeName(Objek) = "TextBox" Then
            Objek.Text = ""
        End If
    Next
    AturKontrol
End Sub


Private Sub cmbJudulBuku_Click()
Kalimat = "JudulBuku = '" & cmbJudulBuku.Text & "'"
AdodcPeminjamanBuku.Refresh
    With AdodcPeminjamanBuku.Recordset
        .Find Kalimat
        If Not .EOF Then
            textKodeBuku.Text = .Fields(6).Value
            textKategori.Text = .Fields(7).Value
            textPengarang.Text = .Fields(8).Value
            textPenerbit.Text = .Fields(9).Value
            textTahunTerbit.Text = .Fields(10).Value
            textCetakanKe.Text = .Fields(11).Value
            
            textHariPinjam.Text = .Fields(12).Value
            textTanggalPinjam.Text = .Fields(13).Value
            textBulanPinjam.Text = .Fields(14).Value
            textTahunPinjam.Text = .Fields(15).Value
            textJumlah.Text = .Fields(16).Value
            textLamaPinjam.Text = .Fields(17).Value
            textKeterangan.Text = .Fields(19).Value
            textAdminSaatPinjam.Text = .Fields(20).Value
            textWaktuPeminjaman.Text = .Fields(26).Value & "-" & .Fields(22).Value & "-" & .Fields(23).Value & "-" & .Fields(24).Value
            
            textNIA.Text = .Fields(0).Value
            textNamaPeminjam.Text = .Fields(1).Value
            textAlamatPeminjam.Text = .Fields(2).Value
            textNoTelpPeminjam.Text = .Fields(3).Value
            textStatusPendidikan.Text = .Fields(4).Value
        End If
    End With
End Sub


Private Sub cmCari_Click()
If AdodcUtama.Recordset.RecordCount = 0 Then
    MsgBox "Tidak ada data yang akan dicari", vbExclamation + vbOKOnly, ""
Else
    With FormCariDataPengembalianBuku
        .Show
        .SetFocus
    End With
End If
End Sub

Private Sub cmEdit_Click()
If AdodcUtama.Recordset.RecordCount = 0 Then
    MsgBox "Tidak ada data yang dapat diedit", vbExclamation + vbOKOnly, ""
Else
    Select Case cmEdit.Caption
    Case "&Edit"
        cmEdit.Caption = "&Batal"
        With Me
            .cmbJudulBuku.Text = .AdodcUtama.Recordset.Fields(0).Value
            .textKodeBuku.Text = .AdodcUtama.Recordset.Fields(1).Value
            .textKategori.Text = .AdodcUtama.Recordset.Fields(2).Value
            .textPengarang.Text = .AdodcUtama.Recordset.Fields(3).Value
            .textPenerbit.Text = .AdodcUtama.Recordset.Fields(4).Value
            .textTahunTerbit.Text = .AdodcUtama.Recordset.Fields(5).Value
            .textCetakanKe.Text = .AdodcUtama.Recordset.Fields(6).Value
            .textHariPinjam.Text = .AdodcUtama.Recordset.Fields(7).Value
            .textTanggalPinjam.Text = .AdodcUtama.Recordset.Fields(8).Value
            .textBulanPinjam.Text = .AdodcUtama.Recordset.Fields(9).Value
            .textTahunPinjam.Text = .AdodcUtama.Recordset.Fields(10).Value
            .textJumlah.Text = .AdodcUtama.Recordset.Fields(11).Value
            .textLamaPinjam.Text = .AdodcUtama.Recordset.Fields(12).Value
            .textKeterangan.Text = .AdodcUtama.Recordset.Fields(13).Value
            .textAdminSaatPinjam.Text = .AdodcUtama.Recordset.Fields(14).Value
            .textWaktuPeminjaman.Text = .AdodcUtama.Recordset.Fields(15).Value
            .textNIA.Text = .AdodcUtama.Recordset.Fields(16).Value
            .textNamaPeminjam.Text = .AdodcUtama.Recordset.Fields(17).Value
            .textAlamatPeminjam.Text = .AdodcUtama.Recordset.Fields(18).Value
            .textNoTelpPeminjam.Text = .AdodcUtama.Recordset.Fields(19).Value
            .textStatusPendidikan.Text = .AdodcUtama.Recordset.Fields(20).Value
            .textNamaAdmin.Text = .AdodcUtama.Recordset.Fields(21).Value
            .textBagian.Text = .AdodcUtama.Recordset.Fields(22).Value
            .textTanggalPengembalian.Text = .AdodcUtama.Recordset.Fields(23).Value
            .textBulanPengembalian.Text = .AdodcUtama.Recordset.Fields(24).Value
            .textTahunPengembalian.Text = .AdodcUtama.Recordset.Fields(25).Value
            .textJam.Text = .AdodcUtama.Recordset.Fields(26).Value
            .textMenit.Text = .AdodcUtama.Recordset.Fields(27).Value
            .textDetik.Text = .AdodcUtama.Recordset.Fields(28).Value
            .cmbHari.Text = .AdodcUtama.Recordset.Fields(29).Value
        End With
        cmSimpan.Caption = "&Perbarui"
    Case "&Batal"
        cmEdit.Caption = "&Edit"
        cmSimpan.Caption = "&Simpan"
        cmReset_Click
    End Select
End If
End Sub

Private Sub cmFilter_Click()
    If AdodcUtama.Recordset.RecordCount = 0 Then
        MsgBox "Tidak ada data yang akan difilter!", vbExclamation + vbOKOnly, ""
    Else
        With FormFilterPengembalianBuku
            .Show
            .SetFocus
        End With
    End If
End Sub

Private Sub cmHapus_Click()
If AdodcUtama.Recordset.RecordCount = 0 Then
    MsgBox "Tidak ada data yang dapat dihapus", vbExclamation + vbOKOnly, ""
Else
    X = MsgBox("Apakah Anda yakin ingin menghapus data ini?", vbQuestion + vbYesNo, "Hapus data?")
    If X = vbYes Then
        With AdodcUtama
            .Recordset.Delete
            .Refresh
        End With
    End If
End If
End Sub

Private Sub cmRefresh_Click()
    Form_Load
    AturKontrol
    cmEdit.Enabled = True
    cmCari.Enabled = True
    cmSorot.Enabled = True
    cmFilter.Enabled = True
    cmHapus.Enabled = True
    cmSimpan.Enabled = True
End Sub

Private Sub cmReset_Click()
    KosongkanInput
    AturKontrol
    textNamaAdmin.SetFocus
End Sub

Private Sub cmSimpan_Click()
    If textNamaAdmin.Text = "" Then
        MsgBox "Silahkan isi Nama Anda pada kolom nama Admin", vbExclamation + vbOKOnly, ""
        textNamaAdmin.SetFocus
    ElseIf textBagian.Text = "" Then
        MsgBox "Silahkan isi nama Bagian Anda (sebagai Admin)", vbExclamation + vbOKOnly, ""
        textBagian.SetFocus
    Else
        Select Case cmSimpan.Caption
        Case "&Simpan"
            X = MsgBox("Apakah Anda yakin ingin menyimpan data ini?", vbQuestion + vbYesNo, "Konfirmasi")
            If X = vbYes Then
                With AdodcUtama
                    .Recordset.AddNew
                    .Recordset.Fields(0).Value = cmbJudulBuku.Text
                    .Recordset.Fields(1).Value = textKodeBuku.Text
                    .Recordset.Fields(2).Value = textKategori.Text
                    .Recordset.Fields(3).Value = textPengarang.Text
                    .Recordset.Fields(4).Value = textPenerbit.Text
                    .Recordset.Fields(5).Value = textTahunTerbit.Text
                    .Recordset.Fields(6).Value = textCetakanKe.Text
                    .Recordset.Fields(7).Value = textHariPinjam.Text
                    .Recordset.Fields(8).Value = textTanggalPinjam.Text
                    .Recordset.Fields(9).Value = textBulanPinjam.Text
                    .Recordset.Fields(10).Value = textTahunPinjam.Text
                    .Recordset.Fields(11).Value = textJumlah.Text
                    .Recordset.Fields(12).Value = textLamaPinjam.Text
                    .Recordset.Fields(13).Value = textKeterangan.Text
                    .Recordset.Fields(14).Value = textAdminSaatPinjam.Text
                    .Recordset.Fields(15).Value = textWaktuPeminjaman.Text
                    .Recordset.Fields(16).Value = textNIA.Text
                    .Recordset.Fields(17).Value = textNamaPeminjam.Text
                    .Recordset.Fields(18).Value = textAlamatPeminjam.Text
                    .Recordset.Fields(19).Value = textNoTelpPeminjam.Text
                    .Recordset.Fields(20).Value = textStatusPendidikan.Text
                    .Recordset.Fields(21).Value = textNamaAdmin.Text
                    .Recordset.Fields(22).Value = textBagian.Text
                    .Recordset.Fields(23).Value = textTanggalPengembalian.Text
                    .Recordset.Fields(24).Value = textBulanPengembalian.Text
                    .Recordset.Fields(25).Value = textTahunPengembalian.Text
                    .Recordset.Fields(26).Value = textDetik.Text
                    .Recordset.Fields(27).Value = textMenit.Text
                    .Recordset.Fields(28).Value = textJam.Text
                    .Recordset.Fields(29).Value = cmbHari.Text
                    .Recordset.Update
                    .Refresh
                End With
                KosongkanInput
            End If
        Case "&Perbarui"
            X = MsgBox("Apakah Anda yakin ingin Memperbarui data ini?", vbQuestion + vbYesNo, "Konfirmasi")
            If X = vbYes Then
                With AdodcUtama
                    .Recordset.Delete
                    .Recordset.AddNew
                    .Recordset.Fields(0).Value = cmbJudulBuku.Text
                    .Recordset.Fields(1).Value = textKodeBuku.Text
                    .Recordset.Fields(2).Value = textKategori.Text
                    .Recordset.Fields(3).Value = textPengarang.Text
                    .Recordset.Fields(4).Value = textPenerbit.Text
                    .Recordset.Fields(5).Value = textTahunTerbit.Text
                    .Recordset.Fields(6).Value = textCetakanKe.Text
                    .Recordset.Fields(7).Value = textHariPinjam.Text
                    .Recordset.Fields(8).Value = textTanggalPinjam.Text
                    .Recordset.Fields(9).Value = textBulanPinjam.Text
                    .Recordset.Fields(10).Value = textTahunPinjam.Text
                    .Recordset.Fields(11).Value = textJumlah.Text
                    .Recordset.Fields(12).Value = textLamaPinjam.Text
                    .Recordset.Fields(13).Value = textKeterangan.Text
                    .Recordset.Fields(14).Value = textAdminSaatPinjam.Text
                    .Recordset.Fields(15).Value = textWaktuPeminjaman.Text
                    .Recordset.Fields(16).Value = textNIA.Text
                    .Recordset.Fields(17).Value = textNamaPeminjam.Text
                    .Recordset.Fields(18).Value = textAlamatPeminjam.Text
                    .Recordset.Fields(19).Value = textNoTelpPeminjam.Text
                    .Recordset.Fields(20).Value = textStatusPendidikan.Text
                    .Recordset.Fields(21).Value = textNamaAdmin.Text
                    .Recordset.Fields(22).Value = textBagian.Text
                    .Recordset.Fields(23).Value = textTanggalPengembalian.Text
                    .Recordset.Fields(24).Value = textBulanPengembalian.Text
                    .Recordset.Fields(25).Value = textTahunPengembalian.Text
                    .Recordset.Fields(26).Value = textDetik.Text
                    .Recordset.Fields(27).Value = textMenit.Text
                    .Recordset.Fields(28).Value = textJam.Text
                    .Recordset.Fields(29).Value = cmbHari.Text
                    .Recordset.Update
                    .Refresh
                End With
                KosongkanInput
                cmSimpan.Caption = "&Simpan"
                cmEdit.Caption = "&Edit"
            End If
        End Select
    End If
End Sub

Private Sub cmSorot_Click()
    If AdodcUtama.Recordset.RecordCount = 0 Then
        MsgBox "Tidak ada data yang akan disorot!", vbExclamation + vbOKOnly, ""
    Else
        With FormSorotPengembalianBuku
            .Show
            .SetFocus
        End With
    End If
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub

Private Sub TimerWaktuPengembalian_Timer()
    textTanggalPengembalian.Text = Day(Date)
    textBulanPengembalian.Text = Month(Date)
    textTahunPengembalian = Year(Date)
    textJam.Text = Hour(Time)
    textMenit.Text = Minute(Time)
    textDetik.Text = Second(Time)
End Sub




VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.ocx"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.ocx"
Begin VB.Form FormPeminjamanBuku 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Peminjaman Buku"
   ClientHeight    =   8760
   ClientLeft      =   5235
   ClientTop       =   1725
   ClientWidth     =   10905
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormTransaksi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   10905
   Begin VB.CommandButton cmSimpan 
      Caption         =   "S&impan"
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   2640
      Width           =   1095
   End
   Begin MSACAL.Calendar Kalender 
      Height          =   2895
      Left            =   4080
      TabIndex        =   60
      Top             =   8760
      Visible         =   0   'False
      Width           =   3855
      _Version        =   524288
      _ExtentX        =   6800
      _ExtentY        =   5106
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2013
      Month           =   3
      Day             =   25
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdodcBuku 
      Height          =   330
      Left            =   240
      Top             =   9000
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
      Height          =   2295
      Left            =   120
      TabIndex        =   40
      Top             =   6360
      Width           =   4695
      Begin VB.Timer TimerWaktuPeminjaman 
         Interval        =   10
         Left            =   4200
         Top             =   1680
      End
      Begin VB.TextBox textBagian 
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         TabIndex        =   52
         Top             =   780
         Width           =   1605
      End
      Begin VB.ComboBox cmbSatuanWaktu 
         Height          =   390
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox textDetik 
         Alignment       =   2  'Center
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   3000
         TabIndex        =   50
         Top             =   1200
         Width           =   645
      End
      Begin VB.TextBox textMenit 
         Alignment       =   2  'Center
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   2160
         TabIndex        =   48
         Top             =   1200
         Width           =   645
      End
      Begin VB.TextBox textNamaAdmin 
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         TabIndex        =   43
         Top             =   360
         Width           =   3075
      End
      Begin VB.TextBox textJam 
         Alignment       =   2  'Center
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         TabIndex        =   42
         Top             =   1200
         Width           =   525
      End
      Begin VB.ComboBox cmbHari 
         Height          =   390
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   1620
         Width           =   2175
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Bagian"
         Height          =   270
         Left            =   120
         TabIndex        =   53
         Top             =   780
         Width           =   420
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2880
         TabIndex        =   49
         Top             =   1200
         Width           =   45
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2040
         TabIndex        =   47
         Top             =   1200
         Width           =   45
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Admin"
         Height          =   270
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Waktu Peminjaman"
         Height          =   270
         Left            =   120
         TabIndex        =   45
         Top             =   1200
         Width           =   1230
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Hari"
         Height          =   270
         Left            =   120
         TabIndex        =   44
         Top             =   1620
         Width           =   270
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2535
      Left            =   120
      TabIndex        =   38
      Top             =   0
      Width           =   10695
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2175
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   3836
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
   End
   Begin VB.CommandButton cmEdit 
      Caption         =   "&Edit"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmCari 
      Caption         =   "&Cari"
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmSorot 
      Caption         =   "&Sorot"
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmFilter 
      Caption         =   "&Filter"
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmHapus 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmRefresh 
      Caption         =   "&Refresh"
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data Peminjam"
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
      Height          =   3015
      Left            =   120
      TabIndex        =   23
      Top             =   3240
      Width           =   4695
      Begin VB.TextBox textNIA 
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         TabIndex        =   61
         Top             =   240
         Width           =   2925
      End
      Begin VB.ComboBox cmbStatusPendidikan 
         Height          =   390
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   2475
         Width           =   1935
      End
      Begin VB.TextBox textNoTelpPeminjam 
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         TabIndex        =   28
         Top             =   2040
         Width           =   2925
      End
      Begin VB.TextBox textAlamatPeminjam 
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   870
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   1140
         Width           =   3165
      End
      Begin VB.TextBox textNamaPeminjam 
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         TabIndex        =   24
         Top             =   720
         Width           =   2925
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "NIA"
         Height          =   270
         Left            =   120
         TabIndex        =   62
         Top             =   240
         Width           =   210
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Status Pendidikan"
         Height          =   270
         Left            =   120
         TabIndex        =   30
         Top             =   2475
         Width           =   1125
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No. Telp Peminjam"
         Height          =   270
         Left            =   120
         TabIndex        =   29
         Top             =   2040
         Width           =   1185
      End
      Begin VB.Label labelalamatpeminjam 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat Peminjam"
         Height          =   270
         Left            =   120
         TabIndex        =   27
         Top             =   1155
         Width           =   1095
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Peminjam"
         Height          =   270
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   1020
      End
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
      Height          =   5415
      Left            =   4920
      TabIndex        =   2
      Top             =   3240
      Width           =   5895
      Begin VB.TextBox textKategori 
         DataField       =   "Judul"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         TabIndex        =   59
         Top             =   1080
         Width           =   2475
      End
      Begin VB.ComboBox cmbSatuanTempo 
         Height          =   390
         ItemData        =   "FormTransaksi.frx":000C
         Left            =   2760
         List            =   "FormTransaksi.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   4500
         Width           =   1635
      End
      Begin VB.TextBox textLamaPinjam 
         Alignment       =   2  'Center
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         TabIndex        =   56
         Top             =   4500
         Width           =   1245
      End
      Begin VB.ComboBox cmbHariPinjam 
         Height          =   390
         ItemData        =   "FormTransaksi.frx":0010
         Left            =   1440
         List            =   "FormTransaksi.frx":0012
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   3660
         Width           =   1275
      End
      Begin VB.ComboBox cmbJudulBuku 
         Height          =   390
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   240
         Width           =   4335
      End
      Begin VB.ComboBox cmbTahunPinjam 
         Height          =   390
         ItemData        =   "FormTransaksi.frx":0014
         Left            =   4440
         List            =   "FormTransaksi.frx":0016
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3660
         Width           =   1395
      End
      Begin VB.ComboBox cmbBulanPinjam 
         Height          =   390
         ItemData        =   "FormTransaksi.frx":0018
         Left            =   3600
         List            =   "FormTransaksi.frx":001A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3660
         Width           =   795
      End
      Begin VB.ComboBox cmbTanggalPinjam 
         Height          =   390
         ItemData        =   "FormTransaksi.frx":001C
         Left            =   2760
         List            =   "FormTransaksi.frx":001E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   3660
         Width           =   795
      End
      Begin VB.TextBox textKodeBuku 
         DataField       =   "Judul"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         TabIndex        =   9
         Top             =   660
         Width           =   2115
      End
      Begin VB.TextBox textPenerbit 
         DataField       =   "Penerbit"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         TabIndex        =   8
         Top             =   2340
         Width           =   4155
      End
      Begin VB.TextBox textTahunTerbit 
         Alignment       =   2  'Center
         DataField       =   "TahunTerbit"
         DataSource      =   "adcBuku"
         Height          =   450
         Left            =   1440
         TabIndex        =   7
         Top             =   2760
         Width           =   2085
      End
      Begin VB.TextBox textCetakanKe 
         Alignment       =   2  'Center
         DataField       =   "Cetakan"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         TabIndex        =   6
         Top             =   3240
         Width           =   2085
      End
      Begin VB.TextBox textJumlah 
         Alignment       =   1  'Right Justify
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         TabIndex        =   5
         Top             =   4080
         Width           =   2085
      End
      Begin VB.TextBox textKeterangan 
         DataField       =   "Keterangan"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   4920
         Width           =   4395
      End
      Begin VB.TextBox textPengarang 
         DataField       =   "Pengarang"
         DataSource      =   "adcBuku"
         Height          =   790
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1500
         Width           =   4395
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Lama Pinjam"
         Height          =   270
         Left            =   120
         TabIndex        =   57
         Top             =   4500
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Buku"
         Height          =   270
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Penerbit"
         Height          =   270
         Left            =   120
         TabIndex        =   21
         Top             =   2400
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tahun Terbit"
         Height          =   270
         Left            =   120
         TabIndex        =   20
         Top             =   2820
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Pinjam"
         Height          =   270
         Left            =   120
         TabIndex        =   19
         Top             =   3720
         Width           =   960
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah"
         Height          =   270
         Left            =   120
         TabIndex        =   18
         Top             =   4100
         Width           =   465
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         Height          =   270
         Left            =   120
         TabIndex        =   17
         Top             =   4920
         Width           =   735
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Pengarang"
         Height          =   270
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   690
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Kategori"
         Height          =   270
         Left            =   120
         TabIndex        =   15
         Top             =   1140
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Judul Buku"
         Height          =   270
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cetakan ke"
         Height          =   270
         Left            =   120
         TabIndex        =   13
         Top             =   3260
         Width           =   705
      End
   End
   Begin VB.CommandButton cmTutup 
      Caption         =   "&Tutup"
      Height          =   495
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmBaru 
      Caption         =   "&Baru"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc AdodcUtama 
      Height          =   330
      Left            =   1440
      Top             =   9000
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
End
Attribute VB_Name = "FormPeminjamanBuku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    NyambunggUtama
    With AdodcUtama
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * from tbPeminjamanBuku order by NamaPeminjam asc;"
        Set DataGrid1.DataSource = AdodcUtama
        .Refresh
    End With
    With AdodcBuku
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * from tbBuku order by JudulBuku asc;"
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
    With cmbBulanPinjam
        .Clear
        .AddItem "01", 0
        .AddItem "02", 1
        .AddItem "03", 2
        .AddItem "04", 3
        .AddItem "05", 4
        .AddItem "06", 5
        .AddItem "07", 6
        .AddItem "08", 7
        .AddItem "09", 8
        .AddItem "10", 9
        .AddItem "11", 10
        .AddItem "12", 11
        .ListIndex = 0
    End With
    cmbTahunPinjam.Clear
    For X = 1800 To 3000
        cmbTahunPinjam.AddItem X
    Next
    cmbTahunPinjam.Text = Year(Date)
    With Me
        .cmbStatusPendidikan.Clear
        .cmbStatusPendidikan.AddItem "TK", 0
        .cmbStatusPendidikan.AddItem "SD", 1
        .cmbStatusPendidikan.AddItem "SMP", 2
        .cmbStatusPendidikan.AddItem "SMA", 3
        .cmbStatusPendidikan.AddItem "Diploma 1", 4
        .cmbStatusPendidikan.AddItem "Diploma 3", 5
        .cmbStatusPendidikan.AddItem "Strata 1", 6
        .cmbStatusPendidikan.AddItem "Strata 2", 7
        .cmbStatusPendidikan.AddItem "Strata 3", 8
        .cmbStatusPendidikan.AddItem "Umum", 9
        .cmbStatusPendidikan.ListIndex = 9
        .cmbSatuanWaktu.Clear
        .cmbSatuanWaktu.AddItem "AM", 0
        .cmbSatuanWaktu.AddItem "PM", 1
        .cmbSatuanWaktu.ListIndex = 0
        .cmbHariPinjam.Clear
        .cmbHariPinjam.AddItem "Senin", 0
        .cmbHariPinjam.AddItem "Selasa", 1
        .cmbHariPinjam.AddItem "Rabu", 2
        .cmbHariPinjam.AddItem "Kamis", 3
        .cmbHariPinjam.AddItem "Jumat", 4
        .cmbHariPinjam.AddItem "Sabtu", 5
        .cmbHariPinjam.AddItem "Minggu", 6
        .cmbHariPinjam.ListIndex = 0
        .cmbHari.Clear
        .cmbHari.AddItem "Senin", 0
        .cmbHari.AddItem "Selasa", 1
        .cmbHari.AddItem "Rabu", 2
        .cmbHari.AddItem "Kamis", 3
        .cmbHari.AddItem "Jumat", 4
        .cmbHari.AddItem "Sabtu", 5
        .cmbHari.AddItem "Minggu", 6
        .cmbHari.ListIndex = 0
        .cmbSatuanTempo.Clear
        .cmbSatuanTempo.AddItem "Menit", 0
        .cmbSatuanTempo.AddItem "Jam", 1
        .cmbSatuanTempo.AddItem "Hari", 2
        .cmbSatuanTempo.AddItem "Pekan/Minggu", 3
        .cmbSatuanTempo.AddItem "Bulan", 4
        .cmbSatuanTempo.AddItem "Tahun", 5
        .cmbSatuanTempo.ListIndex = 2
        .textKodeBuku.Locked = True
        .textPengarang.Locked = True
        .textKategori.Locked = True
        .textPenerbit.Locked = True
        .textTahunTerbit.Locked = True
        .textCetakanKe.Locked = True
        
        AdodcBuku.Refresh
        cmbJudulBuku.Clear
        Do Until AdodcBuku.Recordset.EOF
            cmbJudulBuku.AddItem AdodcBuku.Recordset.Fields(7).Value, 0
            AdodcBuku.Recordset.MoveNext
        Loop
        AdodcBuku.Refresh
        cmbJudulBuku.ListIndex = 0
            
        
    End With
    NonAktifkanInput
    cmSimpan.Enabled = False
    cmbJudulBuku_Click
    With DataGrid1
        .Columns(0).Width = 1000.906
        .Columns(1).Width = 1379.906
        .Columns(2).Width = 2069.858
        .Columns(3).Width = 1275.024
        .Columns(4).Width = 1305.071
        .Columns(5).Width = 2280.189
        .Columns(6).Width = 840.189
        .Columns(7).Width = 824.882
        .Columns(8).Width = 1739.906
        .Columns(9).Width = 1739.906
        .Columns(10).Width = 989.8583
        .Columns(11).Width = 824.882
        .Columns(12).Width = 840.189
        .Columns(13).Width = 1094.74
        .Columns(14).Width = 959.8111
        .Columns(15).Width = 989.8583
        .Columns(16).Width = 569.7638
        .Columns(17).Width = 975.1182
        .Columns(18).Width = 1019.906
        .Columns(19).Width = 1635.024
        .Columns(20).Width = 2115.213
        .Columns(21).Width = 1275.024
        .Columns(22).Width = 1214.929
        .Columns(23).Width = 1319.811
        .Columns(24).Width = 1275.024
        .Columns(25).Width = 1065.26
        .Columns(26).Width = 1739.906
    End With
End Sub
Sub AktifkanInput()
    For Each Objek In Me
        If TypeName(Objek) = "TextBox" Then
            With Objek
                .BackColor = vbWhite
                .Enabled = True
                .Text = ""
                .MaxLength = 254
            End With
        ElseIf TypeName(Objek) = "ComboBox" Then
            With Objek
                .BackColor = vbWhite
                .Enabled = True
            End With
        End If
    Next
End Sub
Sub NonAktifkanInput()
    For Each Objek In Me
        If TypeName(Objek) = "TextBox" Then
            With Objek
                .BackColor = Me.BackColor
                .Enabled = False
                .Text = ""
                .MaxLength = 254
            End With
        ElseIf TypeName(Objek) = "ComboBox" Then
            With Objek
                .BackColor = Me.BackColor
                .Enabled = False
            End With
        End If
    Next
End Sub

Private Sub cmBaru_Click()
    Select Case cmBaru.Caption
    Case "&Baru"
        DataGrid1.Enabled = False
        AktifkanInput
        cmBaru.Caption = "&Batal"
        cmSimpan.Enabled = True
        cmEdit.Enabled = False
        textNamaPeminjam.SetFocus
    Case "&Batal"
        DataGrid1.Enabled = True
        NonAktifkanInput
        cmBaru.Caption = "&Baru"
        cmSimpan.Enabled = False
        cmEdit.Enabled = True
    End Select
End Sub

Private Sub cmbBulanPinjam_Click()
    cmbTanggalPinjam.Clear
    Select Case cmbBulanPinjam.ListIndex
        Case Is = 0, 2, 4, 6, 7, 9, 11
            For X = 1 To 31
                With cmbTanggalPinjam
                    .AddItem X, 0
                End With
            Next
        Case Is = 1
            If Val(cmbTahunPinjam.Text) Mod 4 Then
                For X = 1 To 29
                    With cmbTanggalPinjam
                        .AddItem X, 0
                    End With
                Next
            Else
                For X = 1 To 28
                    With cmbTanggalPinjam
                        .AddItem X, 0
                    End With
                Next
            End If
        Case Is = 3, 5, 8, 10
            For X = 1 To 30
                With cmbTanggalPinjam
                    .AddItem X, 0
                End With
            Next
    End Select
    cmbTanggalPinjam.Text = "1"
End Sub

Private Sub cmbJudulBuku_Click()
Kalimat = "JudulBuku = '" & cmbJudulBuku.Text & "'"
AdodcBuku.Refresh
    With AdodcBuku.Recordset
        .Find Kalimat
        If Not .EOF Then
            textKodeBuku.Text = .Fields(6).Value
            textKategori.Text = .Fields(8).Value
            textPengarang.Text = .Fields(9).Value
            textPenerbit.Text = .Fields(10).Value
            textTahunTerbit.Text = .Fields(11).Value
            textCetakanKe.Text = .Fields(12).Value
        End If
    End With
End Sub

Private Sub cmCari_Click()
If AdodcUtama.Recordset.RecordCount = 0 Then
    MsgBox "Tidak ada data yang akan dicari", vbExclamation + vbOKOnly, ""
Else
    With FormCariPeminjamanBuku
        .Show
        .SetFocus
    End With
End If
End Sub

Private Sub cmEdit_Click()
If AdodcUtama.Recordset.RecordCount = 0 Then
    MsgBox "Maaf tidak ada data yang dapat diedit", vbExclamation + vbOKOnly, ""
Else
    Select Case cmEdit.Caption
    Case "&Edit"
        cmSimpan.Caption = "&Perbarui"
        cmSimpan.Enabled = True
        cmBaru.Enabled = False
        cmEdit.Caption = "&Batal"
        cmRefresh.Enabled = False
        cmCari.Enabled = False
        cmSorot.Enabled = False
        cmFilter.Enabled = False
        cmHapus.Enabled = False
        AktifkanInput
        textNamaPeminjam.SetFocus
            With AdodcUtama
                textNIA.Text = .Recordset.Fields(0).Value
                textNamaPeminjam.Text = .Recordset.Fields(1).Value
                textAlamatPeminjam.Text = .Recordset.Fields(2).Value
                textNoTelpPeminjam.Text = .Recordset.Fields(3).Value
                cmbStatusPendidikan.Text = .Recordset.Fields(4).Value
                cmbJudulBuku.Text = .Recordset.Fields(5).Value
                textKodeBuku.Text = .Recordset.Fields(6).Value
                textKategori.Text = .Recordset.Fields(7).Value
                textPengarang.Text = .Recordset.Fields(8).Value
                textPenerbit.Text = .Recordset.Fields(9).Value
                textTahunTerbit.Text = .Recordset.Fields(10).Value
                textCetakanKe.Text = .Recordset.Fields(11).Value
                cmbHariPinjam.Text = .Recordset.Fields(12).Value
                cmbTanggalPinjam.Text = .Recordset.Fields(13).Value
                cmbBulanPinjam.Text = .Recordset.Fields(14).Value
                cmbTahunPinjam.Text = .Recordset.Fields(15).Value
                textJumlah.Text = .Recordset.Fields(16).Value
                textLamaPinjam.Text = .Recordset.Fields(17).Value
                cmbSatuanTempo.Text = .Recordset.Fields(18).Value
                textKeterangan.Text = .Recordset.Fields(19).Value
                textNamaAdmin.Text = .Recordset.Fields(20).Value
                textBagian.Text = .Recordset.Fields(21).Value
                textJam.Text = .Recordset.Fields(22).Value
                textMenit.Text = .Recordset.Fields(23).Value
                textDetik.Text = .Recordset.Fields(24).Value
                cmbSatuanWaktu.Text = .Recordset.Fields(25).Value
                cmbHari.Text = .Recordset.Fields(26).Value
            End With
    Case "&Batal"
        cmSimpan.Caption = "&Simpan"
        cmSimpan.Enabled = False
        cmBaru.Enabled = True
        cmEdit.Caption = "&Edit"
        cmRefresh.Enabled = True
        cmCari.Enabled = True
        cmSorot.Enabled = True
        cmFilter.Enabled = True
        cmHapus.Enabled = True
        NonAktifkanInput
    End Select
End If
End Sub

Private Sub cmFilter_Click()
If AdodcUtama.Recordset.RecordCount = 0 Then
    MsgBox "Tidak ada data yang akan difilter", vbExclamation + vbOKOnly, ""
Else
    With formFilterPeminjamanBuku
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
    AturKontrol
    cmEdit.Enabled = True
    cmCari.Enabled = True
    cmSorot.Enabled = True
    cmFilter.Enabled = True
    cmHapus.Enabled = True
End Sub

Private Sub cmSimpan_Click()
    If textNIA.Text = "" Then
        MsgBox "Silahkan isi Nomor Induk Anggota!", vbExclamation + vbOKOnly, ""
        textNIA.SetFocus
    ElseIf textNamaPeminjam.Text = "" Then
        MsgBox "Silahkan isi nama peminjam", vbExclamation + vbOKOnly, ""
        textNamaPeminjam.SetFocus
    ElseIf textAlamatPeminjam.Text = "" Then
        MsgBox "Silahkan isi Alamat peminjam", vbExclamation + vbOKOnly, ""
        textAlamatPeminjam.SetFocus
    ElseIf textNoTelpPeminjam.Text = "" Then
        MsgBox "Silahkan isi Nomor Telepon peminjam", vbExclamation + vbOKOnly, ""
        textNoTelpPeminjam.SetFocus
    ElseIf textNamaAdmin.Text = "" Then
        MsgBox "Silahkan isi Nama Admin yang menginput data peminjaman ini (Anda).", vbExclamation + vbOKOnly, ""
        textNamaAdmin.SetFocus
    ElseIf textBagian.Text = "" Then
        MsgBox "Silahkan isi Nama Bagian Admin yang menginput data peminjaman ini (Anda).", vbExclamation + vbOKOnly, ""
        textBagian.SetFocus
    ElseIf textJumlah.Text = "" Then
        MsgBox "Silahkan isi jumlah buku yang dipinjamkan.", vbExclamation + vbOKOnly, ""
        textJumlah.SetFocus
    ElseIf textLamaPinjam.Text = "" Then
        MsgBox "berapa lama buku dipinjamkan?", vbExclamation + vbOKOnly, ""
        textLamaPinjam.SetFocus
    ElseIf textKeterangan.Text = "" Then
        MsgBox "Silahkan isi keterangan yang diperlukan", vbExclamation + vbOKOnly, ""
        textKeterangan.SetFocus
    ElseIf textKodeBuku.Text = "" Or textKategori.Text = "" Or textPengarang.Text = "" Or textPenerbit.Text = "" Or textTahunTerbit.Text = "" Or textCetakanKe.Text = "" Then
        MsgBox "Silahkan pilih judul buku yang akan dipinjamkan", vbExclamation + vbOKOnly, ""
        cmbJudulBuku.SetFocus
    Else
        Select Case cmSimpan.Caption
            Case "S&impan"
                X = MsgBox("Anda yakin ingin menambahkan data baru?", vbQuestion + vbYesNo, "Konfirmasi")
                If X = vbYes Then
                    With AdodcUtama
                        .Recordset.AddNew
                        .Recordset.Fields(0).Value = textNIA.Text
                        .Recordset.Fields(1).Value = textNamaPeminjam.Text
                        .Recordset.Fields(2).Value = textAlamatPeminjam.Text
                        .Recordset.Fields(3).Value = textNoTelpPeminjam.Text
                        .Recordset.Fields(4).Value = cmbStatusPendidikan.Text
                        .Recordset.Fields(5).Value = cmbJudulBuku.Text
                        .Recordset.Fields(6).Value = textKodeBuku.Text
                        .Recordset.Fields(7).Value = textKategori.Text
                        .Recordset.Fields(8).Value = textPengarang.Text
                        .Recordset.Fields(9).Value = textPenerbit.Text
                        .Recordset.Fields(10).Value = textTahunTerbit.Text
                        .Recordset.Fields(11).Value = textCetakanKe.Text
                        .Recordset.Fields(12).Value = cmbHariPinjam.Text
                        .Recordset.Fields(13).Value = cmbTanggalPinjam.Text
                        .Recordset.Fields(14).Value = cmbBulanPinjam.Text
                        .Recordset.Fields(15).Value = cmbTahunPinjam.Text
                        .Recordset.Fields(16).Value = textJumlah.Text
                        .Recordset.Fields(17).Value = textLamaPinjam.Text
                        .Recordset.Fields(18).Value = cmbSatuanTempo.Text
                        .Recordset.Fields(19).Value = textKeterangan.Text
                        .Recordset.Fields(20).Value = textNamaAdmin.Text
                        .Recordset.Fields(21).Value = textBagian.Text
                        .Recordset.Fields(22).Value = textJam.Text
                        .Recordset.Fields(23).Value = textMenit.Text
                        .Recordset.Fields(24).Value = textDetik.Text
                        .Recordset.Fields(25).Value = cmbSatuanWaktu.Text
                        .Recordset.Fields(26).Value = cmbHari.Text
                        .Recordset.Update
                        .Refresh
                    End With
                    cmSimpan.Enabled = False
                    cmBaru.Caption = "&Baru"
                    cmEdit.Enabled = True
                    NonAktifkanInput
                End If
            Case "&Perbarui"
                X = MsgBox("Anda yakin ingin memperbarui data ini?", vbQuestion + vbYesNo, "Konfirmasi")
                If X = vbYes Then
                    With AdodcUtama
                        .Recordset.Delete
                        .Recordset.AddNew
                        .Recordset.Fields(0).Value = textNIA.Text
                        .Recordset.Fields(1).Value = textNamaPeminjam.Text
                        .Recordset.Fields(2).Value = textAlamatPeminjam.Text
                        .Recordset.Fields(3).Value = textNoTelpPeminjam.Text
                        .Recordset.Fields(4).Value = cmbStatusPendidikan.Text
                        .Recordset.Fields(5).Value = cmbJudulBuku.Text
                        .Recordset.Fields(6).Value = textKodeBuku.Text
                        .Recordset.Fields(7).Value = textKategori.Text
                        .Recordset.Fields(8).Value = textPengarang.Text
                        .Recordset.Fields(9).Value = textPenerbit.Text
                        .Recordset.Fields(10).Value = textTahunTerbit.Text
                        .Recordset.Fields(11).Value = textCetakanKe.Text
                        .Recordset.Fields(12).Value = cmbHariPinjam.Text
                        .Recordset.Fields(13).Value = cmbTanggalPinjam.Text
                        .Recordset.Fields(14).Value = cmbBulanPinjam.Text
                        .Recordset.Fields(15).Value = cmbTahunPinjam.Text
                        .Recordset.Fields(16).Value = textJumlah.Text
                        .Recordset.Fields(17).Value = textLamaPinjam.Text
                        .Recordset.Fields(18).Value = cmbSatuanTempo.Text
                        .Recordset.Fields(19).Value = textKeterangan.Text
                        .Recordset.Fields(20).Value = textNamaAdmin.Text
                        .Recordset.Fields(21).Value = textBagian.Text
                        .Recordset.Fields(22).Value = textJam.Text
                        .Recordset.Fields(23).Value = textMenit.Text
                        .Recordset.Fields(24).Value = textDetik.Text
                        .Recordset.Fields(25).Value = cmbSatuanWaktu.Text
                        .Recordset.Fields(26).Value = cmbHari.Text
                        .Recordset.Update
                        .Refresh
                    End With
                    cmSimpan.Enabled = False
                    cmSimpan.Caption = "&Simpan"
                    cmBaru.Caption = "&Baru"
                    cmBaru.Enabled = True
                    cmEdit.Enabled = True
                    cmEdit.Caption = "&Edit"
                    NonAktifkanInput
                    cmRefresh.Enabled = True
                    cmCari.Enabled = True
                    cmSorot.Enabled = True
                    cmFilter.Enabled = True
                    cmHapus.Enabled = True
                End If
            End Select
    End If
    DataGrid1.Enabled = True
End Sub

Private Sub cmSorot_Click()
If AdodcUtama.Recordset.RecordCount = 0 Then
    MsgBox "Tidak ada data yang akan disorot", vbExclamation + vbOKOnly, ""
Else
    With FormSorotPeminjamanBuku
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

Private Sub TimerWaktuPeminjaman_Timer()
    textJam.Text = Hour(Time)
    textMenit.Text = Minute(Time)
    textDetik.Text = Second(Time)
End Sub



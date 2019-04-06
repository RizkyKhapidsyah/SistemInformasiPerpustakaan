VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm FORM_UTAMA 
   BackColor       =   &H8000000C&
   Caption         =   "Sistem Informasi Badan Perpustakaan Arsip & Dokumentasi Provinsi Sumatera Utara (Perpustakaan Daerah) - Dwi Pradana"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   4305
   Icon            =   "FORM_UTAMA.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerSplash 
      Interval        =   10
      Left            =   840
      Top             =   1200
   End
   Begin VB.Timer TimerWaktu 
      Interval        =   10
      Left            =   17520
      Top             =   9000
   End
   Begin MSComctlLib.StatusBar StatusBawah 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   2745
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   17639
            MinWidth        =   17639
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4233
            MinWidth        =   4233
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   1164
      ButtonWidth     =   609
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   393216
   End
   Begin VB.Menu menuFile 
      Caption         =   "File"
      Begin VB.Menu menuPegawai 
         Caption         =   "Anggota"
      End
      Begin VB.Menu MenuDataBuku 
         Caption         =   "Data Buku"
      End
      Begin VB.Menu menuDBK 
         Caption         =   "Data Barang Koleksi"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu menuAdmin 
         Caption         =   "Admin"
         Begin VB.Menu menuDaftarAdmin 
            Caption         =   "Daftar Admin"
         End
         Begin VB.Menu sep5 
            Caption         =   "-"
         End
         Begin VB.Menu menuLO 
            Caption         =   "Log Out"
         End
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu menuKeluar 
         Caption         =   "Keluar"
      End
   End
   Begin VB.Menu menuTransaksi 
      Caption         =   "Transaksi"
      Begin VB.Menu menuBuku 
         Caption         =   "Buku"
         Begin VB.Menu menuBM 
            Caption         =   "Buku Masuk"
         End
         Begin VB.Menu sep4 
            Caption         =   "-"
         End
         Begin VB.Menu menuPeminjaman 
            Caption         =   "Peminjaman"
         End
         Begin VB.Menu menuPengembalian 
            Caption         =   "Pengembalian"
         End
      End
      Begin VB.Menu menuBarangKoleksi 
         Caption         =   "Barang Koleksi"
         Begin VB.Menu menuBKM 
            Caption         =   "Barang Koleksi Masuk"
         End
      End
   End
   Begin VB.Menu menuLaporan 
      Caption         =   "Laporan"
      Begin VB.Menu menuLaporanBuku 
         Caption         =   "Buku"
         Begin VB.Menu menuLaporanPengembalianBuku 
            Caption         =   "Laporan Pengembalian Buku"
         End
         Begin VB.Menu menuLaporanPeminjamanBuku 
            Caption         =   "Laporan Peminjaman Buku"
         End
      End
      Begin VB.Menu menuLBK 
         Caption         =   "Laporan Barang Koleksi"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu menuLaporanDaftarAnggota 
         Caption         =   "Laporan DaftarAnggota"
      End
   End
   Begin VB.Menu menuTools 
      Caption         =   "Tools"
      Begin VB.Menu menuPengaturan 
         Caption         =   "Pengaturan"
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "Help"
      Begin VB.Menu menuSplash 
         Caption         =   "Splash"
      End
      Begin VB.Menu sep9 
         Caption         =   "-"
      End
      Begin VB.Menu menuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "FORM_UTAMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
With StatusBawah.Panels
    .Item(1).Alignment = sbrLeft
    .Item(2).Alignment = sbrCenter
    .Item(3).Alignment = sbrCenter
    .Item(4).Alignment = sbrCenter
    .Item(5).Alignment = sbrLeft
    .Item(1).ToolTipText = .Item(1).Text
    .Item(2).ToolTipText = "Admin"
    .Item(3).ToolTipText = "Tanggal Hari Ini"
    .Item(4).ToolTipText = "Waktu Saat Ini"
    .Item(5).ToolTipText = "Database System"
End With
End Sub

Private Sub MDIForm_Load()
    AturKontrol
    WindowState = vbMaximized
End Sub

Private Sub menuAbout_Click()
    MsgBox "Sistem Informasi Badan Perpustakaan Arsip & Dokumentasi Provinsi Sumatera Utara (Perpustakaan Daerah) - Dwi Pradana", vbInformation + vbOKOnly, "About"
End Sub

Private Sub menuBKM_Click()
    With FormBarangKoleksiMasuk
        .Show
        .SetFocus
    End With
End Sub

Private Sub menuBM_Click()
With FormBukuMasuk
    .Show
    .SetFocus
End With
End Sub

Private Sub menuDaftarAdmin_Click()
With FormAutorisasi
    .Caption = "Autentikasi - (Admin - @" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & ")"
    .textAutentikasiSembunyi.Text = FORM_UTAMA.StatusBawah.Panels.Item(2).Text
    .Caption = "Autorisasi - (Admin@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & ")"
    .LabelJenisAksi.Caption = "TampilkanFormDaftarAdmin"
    .Show vbModal
End With
End Sub

Private Sub MenuDataBuku_Click()
With FormManageBukuMasuk
    .Show
    .SetFocus
End With
End Sub

Private Sub menuDBK_Click()
With FormManageBarangKoleksiMasuk
    .Show
    .SetFocus
End With
End Sub

Private Sub menuKeluar_Click()
    X = MsgBox("Apakah Anda yakin ingin keluar?", vbQuestion + vbYesNo, "Keluar?")
    If X = vbYes Then
        End
    End If
End Sub

Private Sub menuLaporanDaftarAnggota_Click()
With FormLaporanDaftarAnggota
    .Show
    .SetFocus
End With
End Sub

Private Sub menuLaporanPeminjamanBuku_Click()
    With FormLaporanPeminjamanBuku
        .Show
        .SetFocus
    End With
End Sub

Private Sub menuLaporanPengembalianBuku_Click()
With FormLaporanPengembalianBuku
    .Show
    .SetFocus
End With
End Sub

Private Sub menuLBK_Click()
With FormLaporanBarangKoleksi
    .Show
    .SetFocus
End With
End Sub

Private Sub menuLO_Click()
    X = MsgBox("Anda yakin ingin logout dari akun Anda?", vbQuestion + vbYesNo, "LogOut?")
    If X = vbYes Then
        Unload FORM_UTAMA
        With FormLogin
            .Show
            .SetFocus
        End With
    End If
End Sub

Private Sub menuPegawai_Click()
    With FormAnggota
        .Show
        .SetFocus
    End With
End Sub

Private Sub menuPeminjaman_Click()
    With FormPeminjamanBuku
        .Show
        .SetFocus
    End With
End Sub

Private Sub menuPengaturan_Click()
    FormPengaturan.Show vbModal, Me
End Sub

Private Sub menuPengembalian_Click()
    With FormPengembalianBuku
        .Show
        .SetFocus
    End With
End Sub



Private Sub menuSplash_Click()
    FormSplash.Show vbModal, Me
End Sub

Private Sub TimerWaktu_Timer()
Select Case Month(Date)
    Case Is = 1
        Kalimat = "Januari"
    Case Is = 2
        Kalimat = "Februari"
    Case Is = 3
        Kalimat = "Maret"
    Case Is = 4
        Kalimat = "April"
    Case Is = 5
        Kalimat = "Mei"
    Case Is = 6
        Kalimat = "Juni"
    Case Is = 7
        Kalimat = "Juli"
    Case Is = 8
        Kalimat = "Agustus"
    Case Is = 9
        Kalimat = "September"
    Case Is = 10
        Kalimat = "Oktober"
    Case Is = 11
        Kalimat = "November"
    Case Is = 12
        Kalimat = "Desember"
End Select
With StatusBawah.Panels
    .Item(3).Text = Day(Date) & " " & Kalimat & " " & Year(Date)
    .Item(4).Text = Time
End With
End Sub

VERSION 5.00
Begin VB.Form formFilterPeminjamanBuku 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filter Data"
   ClientHeight    =   1950
   ClientLeft      =   6795
   ClientTop       =   5445
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "formFilterPeminjamanBuku.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   6255
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   6015
      Begin VB.ComboBox cmbFilterBerdasarkan 
         Height          =   390
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   3975
      End
      Begin VB.ComboBox cmbMode 
         Height          =   390
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   2535
      End
      Begin VB.CommandButton cmFilter 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Filter"
         Height          =   375
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filter Berdasarkan :"
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1290
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dengan Mode :"
         Height          =   270
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   915
      End
   End
   Begin VB.CommandButton cmBatal 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Batal"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1440
      Width           =   1335
   End
End
Attribute VB_Name = "formFilterPeminjamanBuku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
With Me
    .cmbFilterBerdasarkan.Clear
    .cmbFilterBerdasarkan.AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(0).Name, 0
    .cmbFilterBerdasarkan.AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(1).Name, 1
    .cmbFilterBerdasarkan.AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(2).Name, 2
    .cmbFilterBerdasarkan.AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(3).Name, 3
    .cmbFilterBerdasarkan.AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(4).Name, 4
    .cmbFilterBerdasarkan.AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(5).Name, 5
    .cmbFilterBerdasarkan.AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(6).Name, 6
    .cmbFilterBerdasarkan.AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(7).Name, 7
    .cmbFilterBerdasarkan.AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(8).Name, 8
    .cmbFilterBerdasarkan.AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(9).Name, 9
    .cmbFilterBerdasarkan.AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(10).Name, 10
    .cmbFilterBerdasarkan.AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(11).Name, 11
    .cmbFilterBerdasarkan.AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(12).Name, 12
    .cmbFilterBerdasarkan.AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(13).Name, 13
    .cmbFilterBerdasarkan.AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(14).Name, 14
    .cmbFilterBerdasarkan.AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(15).Name, 15
    .cmbFilterBerdasarkan.AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(16).Name, 16
    .cmbFilterBerdasarkan.AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(17).Name, 17
    .cmbFilterBerdasarkan.AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(18).Name, 18
    .cmbFilterBerdasarkan.AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(19).Name, 19
    .cmbFilterBerdasarkan.AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(20).Name, 20
    .cmbFilterBerdasarkan.AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(21).Name, 21
    .cmbFilterBerdasarkan.AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(22).Name, 22
    .cmbFilterBerdasarkan.AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(23).Name, 23
    .cmbFilterBerdasarkan.AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(24).Name, 24
    .cmbFilterBerdasarkan.AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(25).Name, 25
    .cmbFilterBerdasarkan.AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(26).Name, 26
    .cmbFilterBerdasarkan.ListIndex = 0
    .cmbMode.Clear
    .cmbMode.AddItem "Asc", 0
    .cmbMode.AddItem "Desc", 1
    .cmbMode.ListIndex = 0
End With
End Sub

Private Sub cmBatal_Click()
    Unload Me
End Sub

Private Sub cmFilter_Click()
If cmbMode.ListIndex = 0 Then
    With FormPeminjamanBuku
        .AdodcUtama.Refresh
            Select Case cmbFilterBerdasarkan.ListIndex
            Case Is = 0
                .AdodcUtama.RecordSource = "Select NIA from tbPeminjamanBuku order by NamaPeminjam asc;"
            Case Is = 1
                .AdodcUtama.RecordSource = "Select Nama_Peminjam from tbPeminjamanBuku order by NamaPeminjam asc;"
            Case Is = 2
                .AdodcUtama.RecordSource = "Select Alamat_Peminjam from tbPeminjamanBuku order by AlamatPeminjam asc;"
            Case Is = 3
                .AdodcUtama.RecordSource = "Select No_Telp_Peminjam from tbPeminjamanBuku order by NoTelpPeminjam asc;"
            Case Is = 4
                .AdodcUtama.RecordSource = "Select Status_Pendidikan from tbPeminjamanBuku order by StatusPendidikan asc;"
            Case Is = 5
                .AdodcUtama.RecordSource = "Select Judul_Buku from tbPeminjamanBuku order by JudulBuku asc;"
            Case Is = 6
                .AdodcUtama.RecordSource = "Select Kode_Buku from tbPeminjamanBuku order by KodeBuku asc;"
            Case Is = 7
                .AdodcUtama.RecordSource = "Select Kategori from tbPeminjamanBuku order by Kategori asc;"
            Case Is = 8
                .AdodcUtama.RecordSource = "Select Pengarang from tbPeminjamanBuku order by Pengarang asc;"
            Case Is = 9
                .AdodcUtama.RecordSource = "Select Penerbit from tbPeminjamanBuku order by Penerbit asc;"
            Case Is = 10
                .AdodcUtama.RecordSource = "Select Tahun_Terbit from tbPeminjamanBuku order by TahunTerbit asc;"
            Case Is = 11
                .AdodcUtama.RecordSource = "Select Cetakan_Ke from tbPeminjamanBuku order by CetakanKe asc;"
            Case Is = 12
                .AdodcUtama.RecordSource = "Select Hari_Pinjam from tbPeminjamanBuku order by HariPinjam asc;"
            Case Is = 13
                .AdodcUtama.RecordSource = "Select Tanggal_Pinjam from tbPeminjamanBuku order by TanggalPinjam asc;"
            Case Is = 14
                .AdodcUtama.RecordSource = "Select Bulan_Pinjam from tbPeminjamanBuku order by BulanPinjam asc;"
            Case Is = 15
                .AdodcUtama.RecordSource = "Select Tahun_Pinjam from tbPeminjamanBuku order by TahunPinjam asc;"
            Case Is = 16
                .AdodcUtama.RecordSource = "Select Jumlah from tbPeminjamanBuku order by Jumlah asc;"
            Case Is = 17
                .AdodcUtama.RecordSource = "Select Lama_Pinjam from tbPeminjamanBuku order by LamaPinjam asc;"
            Case Is = 18
                .AdodcUtama.RecordSource = "Select Satuan_Tempo from tbPeminjamanBuku order by SatuanTempo asc;"
            Case Is = 19
                .AdodcUtama.RecordSource = "Select Keterangan from tbPeminjamanBuku order by Keterangan asc;"
            Case Is = 20
                .AdodcUtama.RecordSource = "Select Nama_Admin from tbPeminjamanBuku order by NamaAdmin asc;"
            Case Is = 21
                .AdodcUtama.RecordSource = "Select Bagian from tbPeminjamanBuku order by Bagian asc;"
            Case Is = 22
                .AdodcUtama.RecordSource = "Select Jam_Pinjam from tbPeminjamanBuku order by JamPinjam asc;"
            Case Is = 23
                .AdodcUtama.RecordSource = "Select Menit_Pinjam from tbPeminjamanBuku order by MenitPinjam asc;"
            Case Is = 24
                .AdodcUtama.RecordSource = "Select Detik_Pinjam from tbPeminjamanBuku order by DetikPinjam asc;"
            Case Is = 25
                .AdodcUtama.RecordSource = "Select Satuan_Waktu from tbPeminjamanBuku order by SatuanWaktu asc;"
            Case Is = 26
                .AdodcUtama.RecordSource = "Select Hari_Input_Data from tbPeminjamanBuku order by HariInput_Data asc;"
            End Select
    End With
ElseIf cmbMode.ListIndex = 1 Then
    With FormPeminjamanBuku
        .AdodcUtama.Refresh
            Select Case cmbFilterBerdasarkan.ListIndex
            Case Is = 0
                .AdodcUtama.RecordSource = "Select NIA from tbPeminjamanBuku order by NamaPeminjam desc;"
            Case Is = 1
                .AdodcUtama.RecordSource = "Select Nama_Peminjam from tbPeminjamanBuku order by NamaPeminjam desc;"
            Case Is = 2
                .AdodcUtama.RecordSource = "Select Alamat_Peminjam from tbPeminjamanBuku order by AlamatPeminjam desc;"
            Case Is = 3
                .AdodcUtama.RecordSource = "Select No_Telp_Peminjam from tbPeminjamanBuku order by NoTelpPeminjam desc;"
            Case Is = 4
                .AdodcUtama.RecordSource = "Select Status_Pendidikan from tbPeminjamanBuku order by StatusPendidikan desc;"
            Case Is = 5
                .AdodcUtama.RecordSource = "Select Judul_Buku from tbPeminjamanBuku order by JudulBuku desc;"
            Case Is = 6
                .AdodcUtama.RecordSource = "Select Kode_Buku from tbPeminjamanBuku order by KodeBuku desc;"
            Case Is = 7
                .AdodcUtama.RecordSource = "Select Kategori from tbPeminjamanBuku order by Kategori desc;"
            Case Is = 8
                .AdodcUtama.RecordSource = "Select Pengarang from tbPeminjamanBuku order by Pengarang desc;"
            Case Is = 9
                .AdodcUtama.RecordSource = "Select Penerbit from tbPeminjamanBuku order by Penerbit desc;"
            Case Is = 10
                .AdodcUtama.RecordSource = "Select Tahun_Terbit from tbPeminjamanBuku order by TahunTerbit desc;"
            Case Is = 11
                .AdodcUtama.RecordSource = "Select Cetakan_Ke from tbPeminjamanBuku order by CetakanKe desc;"
            Case Is = 12
                .AdodcUtama.RecordSource = "Select Hari_Pinjam from tbPeminjamanBuku order by HariPinjam desc;"
            Case Is = 13
                .AdodcUtama.RecordSource = "Select Tanggal_Pinjam from tbPeminjamanBuku order by TanggalPinjam desc;"
            Case Is = 14
                .AdodcUtama.RecordSource = "Select Bulan_Pinjam from tbPeminjamanBuku order by BulanPinjam desc;"
            Case Is = 15
                .AdodcUtama.RecordSource = "Select Tahun_Pinjam from tbPeminjamanBuku order by TahunPinjam desc;"
            Case Is = 16
                .AdodcUtama.RecordSource = "Select Jumlah from tbPeminjamanBuku order by Jumlah desc;"
            Case Is = 17
                .AdodcUtama.RecordSource = "Select Lama_Pinjam from tbPeminjamanBuku order by LamaPinjam desc;"
            Case Is = 18
                .AdodcUtama.RecordSource = "Select Satuan_Tempo from tbPeminjamanBuku order by SatuanTempo desc;"
            Case Is = 19
                .AdodcUtama.RecordSource = "Select Keterangan from tbPeminjamanBuku order by Keterangan desc;"
            Case Is = 20
                .AdodcUtama.RecordSource = "Select Nama_Admin from tbPeminjamanBuku order by NamaAdmin desc;"
            Case Is = 21
                .AdodcUtama.RecordSource = "Select Bagian from tbPeminjamanBuku order by Bagian desc;"
            Case Is = 22
                .AdodcUtama.RecordSource = "Select Jam_Pinjam from tbPeminjamanBuku order by JamPinjam desc;"
            Case Is = 23
                .AdodcUtama.RecordSource = "Select Menit_Pinjam from tbPeminjamanBuku order by MenitPinjam desc;"
            Case Is = 24
                .AdodcUtama.RecordSource = "Select Detik_Pinjam from tbPeminjamanBuku order by DetikPinjam desc;"
            Case Is = 25
                .AdodcUtama.RecordSource = "Select Satuan_Waktu from tbPeminjamanBuku order by SatuanWaktu desc;"
            Case Is = 26
                .AdodcUtama.RecordSource = "Select Hari_Input_Data from tbPeminjamanBuku order by HariInput_Data desc;"
            End Select
    End With
End If
    FormPeminjamanBuku.AdodcUtama.Refresh
    cmBatal.Caption = "&Tutup"
    If FormPengaturan.cekTutupFormFilter.Value = Checked Then Me.Hide
    With FormPeminjamanBuku
        .cmEdit.Enabled = False
        .cmCari.Enabled = False
        .cmSorot.Enabled = False
        .cmFilter.Enabled = False
        .cmHapus.Enabled = False
        
    End With
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub

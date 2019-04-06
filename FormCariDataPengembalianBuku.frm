VERSION 5.00
Begin VB.Form FormCariDataPengembalianBuku 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cari Data"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormCariDataPengembalianBuku.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   6000
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   5775
      Begin VB.ComboBox cmbCariDataBerdasarkan 
         Height          =   390
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox textKriteria 
         Height          =   390
         Left            =   2040
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   720
         Width           =   2415
      End
      Begin VB.CommandButton cmCari 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cari"
         Height          =   375
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cari data berdasarkan"
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dengan Kriteria"
         Height          =   270
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmBatal 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Batal"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
   End
End
Attribute VB_Name = "FormCariDataPengembalianBuku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    With textKriteria
        .Text = ""
        .MaxLength = 254
    End With
    With cmbCariDataBerdasarkan
        .Clear
        .AddItem FormPengembalianBuku.AdodcUtama.Recordset.Fields(0).Name, 0
        .AddItem FormPengembalianBuku.AdodcUtama.Recordset.Fields(1).Name, 1
        .AddItem FormPengembalianBuku.AdodcUtama.Recordset.Fields(2).Name, 2
        .AddItem FormPengembalianBuku.AdodcUtama.Recordset.Fields(3).Name, 3
        .AddItem FormPengembalianBuku.AdodcUtama.Recordset.Fields(4).Name, 4
        .AddItem FormPengembalianBuku.AdodcUtama.Recordset.Fields(5).Name, 5
        .AddItem FormPengembalianBuku.AdodcUtama.Recordset.Fields(6).Name, 6
        .AddItem FormPengembalianBuku.AdodcUtama.Recordset.Fields(7).Name, 7
        .AddItem FormPengembalianBuku.AdodcUtama.Recordset.Fields(8).Name, 8
        .AddItem FormPengembalianBuku.AdodcUtama.Recordset.Fields(9).Name, 9
        .AddItem FormPengembalianBuku.AdodcUtama.Recordset.Fields(10).Name, 10
        .AddItem FormPengembalianBuku.AdodcUtama.Recordset.Fields(11).Name, 11
        .AddItem FormPengembalianBuku.AdodcUtama.Recordset.Fields(12).Name, 12
        .AddItem FormPengembalianBuku.AdodcUtama.Recordset.Fields(13).Name, 13
        .AddItem FormPengembalianBuku.AdodcUtama.Recordset.Fields(14).Name, 14
        .AddItem FormPengembalianBuku.AdodcUtama.Recordset.Fields(15).Name, 15
        .AddItem FormPengembalianBuku.AdodcUtama.Recordset.Fields(16).Name, 16
        .AddItem FormPengembalianBuku.AdodcUtama.Recordset.Fields(17).Name, 17
        .AddItem FormPengembalianBuku.AdodcUtama.Recordset.Fields(18).Name, 18
        .AddItem FormPengembalianBuku.AdodcUtama.Recordset.Fields(19).Name, 19
        .AddItem FormPengembalianBuku.AdodcUtama.Recordset.Fields(20).Name, 20
        .AddItem FormPengembalianBuku.AdodcUtama.Recordset.Fields(21).Name, 21
        .AddItem FormPengembalianBuku.AdodcUtama.Recordset.Fields(22).Name, 22
        .AddItem FormPengembalianBuku.AdodcUtama.Recordset.Fields(23).Name, 23
        .AddItem FormPengembalianBuku.AdodcUtama.Recordset.Fields(24).Name, 24
        .AddItem FormPengembalianBuku.AdodcUtama.Recordset.Fields(25).Name, 25
        .AddItem FormPengembalianBuku.AdodcUtama.Recordset.Fields(26).Name, 26
        .AddItem FormPengembalianBuku.AdodcUtama.Recordset.Fields(27).Name, 27
        .ListIndex = 4
    End With
End Sub

Private Sub cmBatal_Click()
    Unload Me
End Sub

Private Sub cmCari_Click()
On Error Resume Next
If textKriteria.Text = "" Then
    MsgBox "Silahkan isi data yang akan dicari!", vbExclamation + vbOKOnly, ""
    textKriteria.SetFocus
Else
    cmBatal.Caption = "&Tutup"
    FormPengembalianBuku.AdodcUtama.Refresh
    With FormPengembalianBuku.AdodcUtama.Recordset
        Select Case cmbCariDataBerdasarkan.ListIndex
        Case Is = 0
            .Find "Judul_Buku = '" & textKriteria.Text & "'"
        Case Is = 1
            .Find "Kode_Buku = '" & textKriteria.Text & "'"
        Case Is = 2
            .Find "Kategori = '" & textKriteria.Text & "'"
        Case Is = 3
            .Find "Pengarang = '" & textKriteria.Text & "'"
        Case Is = 4
            .Find "Penerbit = '" & textKriteria.Text & "'"
        Case Is = 5
            .Find "Tahun_Terbit = '" & textKriteria.Text & "'"
        Case Is = 6
            .Find "Cetakan_Ke = '" & textKriteria.Text & "'"
        Case Is = 7
            .Find "Tanggal_Pinjam = '" & textKriteria.Text & "'"
        Case Is = 8
            .Find "Bulan_Pinjam = '" & textKriteria.Text & "'"
        Case Is = 9
            .Find "Tahun_Pinjam = '" & textKriteria.Text & "'"
        Case Is = 10
            .Find "Jumlah = '" & textKriteria.Text & "'"
        Case Is = 11
            .Find "Lama_Pinjam = '" & textKriteria.Text & "'"
        Case Is = 12
            .Find "Keterangan = '" & textKriteria.Text & "'"
        Case Is = 13
            .Find "Admin_Saat_Pinjam = '" & textKriteria.Text & "'"
        Case Is = 14
            .Find "Waktu_Peminjaman = '" & textKriteria.Text & "'"
        Case Is = 15
            .Find "NIA = '" & textKriteria.Text & "'"
        Case Is = 16
            .Find "Nama_Peminjam = '" & textKriteria.Text & "'"
        Case Is = 17
            .Find "Alamat_Peminjam = '" & textKriteria.Text & "'"
        Case Is = 18
            .Find "No_Telp_Peminjam = '" & textKriteria.Text & "'"
        Case Is = 19
            .Find "Status_Pendidikan = '" & textKriteria.Text & "'"
        Case Is = 20
            .Find "Nama_Admin = '" & textKriteria.Text & "'"
        Case Is = 21
            .Find "Bagian = '" & textKriteria.Text & "'"
        Case Is = 22
            .Find "Tanggal_Pengembalian = '" & textKriteria.Text & "'"
        Case Is = 23
            .Find "Bulan_Pengembalian = '" & textKriteria.Text & "'"
        Case Is = 24
            .Find "Tahun_Pengembalian = '" & textKriteria.Text & "'"
        Case Is = 25
            .Find "Detik_Waktu_Pengembalian = '" & textKriteria.Text & "'"
        Case Is = 26
            .Find "Menit_Waktu_Pengembalian = '" & textKriteria.Text & "'"
        Case Is = 27
            .Find "Jam_Waktu_Pengembalian = '" & textKriteria.Text & "'"
        Case Is = 28
            .Find "Hari = '" & textKriteria.Text & "'"
        End Select
        If .EOF Then
            MsgBox ("Data tidak ditemukan"), vbExclamation + vbOKOnly, ""
        Else
            Set FormPengembalianBuku.DataGrid1.DataSource = FormPengembalianBuku.AdodcUtama.Recordset
            cmBatal.Caption = "&Tutup"
            If FormPengaturan.cekTutupFormCari.Value = Checked Then Me.Hide
        End If
    End With
End If
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub


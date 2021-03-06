VERSION 5.00
Begin VB.Form FormSorotPengembalianBuku 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sorot Data"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6210
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormSorotPengembalianBuku.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   6210
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   6015
      Begin VB.ComboBox cmbSorotDataBerdasarkan 
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   3735
      End
      Begin VB.ComboBox cmbMode 
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton cmSorot 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Sorot"
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sorot Data Berdasarkan"
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dengan Mode"
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   825
      End
   End
   Begin VB.CommandButton cmBatal 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Batal"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
End
Attribute VB_Name = "FormSorotPengembalianBuku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    With cmbSorotDataBerdasarkan
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
    With Me
        .cmbMode.Clear
        .cmbMode.AddItem "Asc", 0
        .cmbMode.AddItem "Desc", 1
        .cmbMode.ListIndex = 0
    End With
End Sub

Private Sub cmBatal_Click()
    Unload Me
End Sub

Private Sub cmSorot_Click()
If cmbMode.ListIndex = 0 Then
    With FormPengembalianBuku
        .AdodcUtama.Refresh
            Select Case cmbSorotDataBerdasarkan.ListIndex
            Case Is = 0
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Judul_Buku asc;"
            Case Is = 1
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Kode_Buku asc;"
            Case Is = 2
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Kategori asc;"
            Case Is = 3
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Pengarang asc;"
            Case Is = 4
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Penerbit asc;"
            Case Is = 5
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Tahun_Terbit asc;"
            Case Is = 6
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Cetakan_Ke asc;"
            Case Is = 7
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Tanggal_Pinjam asc;"
            Case Is = 8
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Bulan_Pinjam asc;"
            Case Is = 9
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Tahun_Pinjam asc;"
            Case Is = 10
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Jumlah asc;"
            Case Is = 11
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Lama_Pinjam asc;"
            Case Is = 12
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Keterangan asc;"
            Case Is = 13
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Admin_Saat_Pinjam asc;"
            Case Is = 14
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Waktu_Peminjaman asc;"
            Case Is = 15
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Nama_Peminjam asc;"
            Case Is = 16
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Alamat_Peminjam asc;"
            Case Is = 17
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by No_Telp_Peminjam asc;"
            Case Is = 18
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Status_Pendidikan asc;"
            Case Is = 19
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Nama_Admin asc;"
            Case Is = 20
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Bagian asc;"
            Case Is = 21
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Tanggal_Pengembalian asc;"
            Case Is = 22
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Bulan_Pengembalian asc;"
            Case Is = 23
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Tahun_Pengembalian asc;"
            Case Is = 24
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Detik_Waktu_Pengembalian asc;"
            Case Is = 25
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Menit_Waktu_Pengembalian asc;"
            Case Is = 26
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Jam_Waktu_Pengembalian asc;"
            Case Is = 27
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Hari asc;"
            End Select
    End With
ElseIf cmbMode.ListIndex = 1 Then
    With FormPengembalianBuku
        .AdodcUtama.Refresh
            Select Case cmbSorotDataBerdasarkan.ListIndex
            Case Is = 0
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Judul_Buku desc;"
            Case Is = 1
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Kode_Buku desc;"
            Case Is = 2
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Kategori desc;"
            Case Is = 3
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Pengarang desc;"
            Case Is = 4
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Penerbit desc;"
            Case Is = 5
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Tahun_Terbit desc;"
            Case Is = 6
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Cetakan_Ke desc;"
            Case Is = 7
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Tanggal_Pinjam desc;"
            Case Is = 8
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Bulan_Pinjam desc;"
            Case Is = 9
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Tahun_Pinjam desc;"
            Case Is = 10
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Jumlah desc;"
            Case Is = 11
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Lama_Pinjam desc;"
            Case Is = 12
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Keterangan desc;"
            Case Is = 13
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Admin_Saat_Pinjam desc;"
            Case Is = 14
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Waktu_Peminjaman desc;"
            Case Is = 15
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Nama_Peminjam desc;"
            Case Is = 16
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Alamat_Peminjam desc;"
            Case Is = 17
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by No_Telp_Peminjam desc;"
            Case Is = 18
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Status_Pendidikan desc;"
            Case Is = 19
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Nama_Admin desc;"
            Case Is = 20
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Bagian desc;"
            Case Is = 21
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Tanggal_Pengembalian desc;"
            Case Is = 22
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Bulan_Pengembalian desc;"
            Case Is = 23
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Tahun_Pengembalian desc;"
            Case Is = 24
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Detik_Waktu_Pengembalian desc;"
            Case Is = 25
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Menit_Waktu_Pengembalian desc;"
            Case Is = 26
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Jam_Waktu_Pengembalian desc;"
            Case Is = 27
                .AdodcUtama.RecordSource = "Select * from tbPeminjamanBuku order by Hari desc;"
        End Select
    End With
End If
    FormPengembalianBuku.AdodcUtama.Refresh
    cmBatal.Caption = "&Tutup"
    If FormPengaturan.cekTutupFormSorot.Value = Checked Then Me.Hide
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub


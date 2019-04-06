VERSION 5.00
Begin VB.Form FormCariPeminjamanBuku 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cari Data"
   ClientHeight    =   1800
   ClientLeft      =   8040
   ClientTop       =   3990
   ClientWidth     =   6030
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormCariPeminjamanBuku.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   6030
   Begin VB.CommandButton cmBatal 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Batal"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.CommandButton cmCari 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cari"
         Height          =   375
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox textKriteria 
         Height          =   390
         Left            =   2040
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   720
         Width           =   2415
      End
      Begin VB.ComboBox cmbCariDataBerdasarkan 
         Height          =   390
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dengan Kriteria"
         Height          =   270
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cari data berdasarkan"
         Height          =   270
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FormCariPeminjamanBuku"
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
        .AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(0).Name, 0
        .AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(1).Name, 1
        .AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(2).Name, 2
        .AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(3).Name, 3
        .AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(4).Name, 4
        .AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(5).Name, 5
        .AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(6).Name, 6
        .AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(7).Name, 7
        .AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(8).Name, 8
        .AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(9).Name, 9
        .AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(10).Name, 10
        .AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(11).Name, 11
        .AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(12).Name, 12
        .AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(13).Name, 13
        .AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(14).Name, 14
        .AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(15).Name, 15
        .AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(16).Name, 16
        .AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(17).Name, 17
        .AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(18).Name, 18
        .AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(19).Name, 19
        .AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(20).Name, 20
        .AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(21).Name, 21
        .AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(22).Name, 22
        .AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(23).Name, 23
        .AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(24).Name, 24
        .AddItem FormPeminjamanBuku.AdodcUtama.Recordset.Fields(25).Name, 25
        .ListIndex = 4
    End With
End Sub

Private Sub cmBatal_Click()
    Unload Me
End Sub

Private Sub cmCari_Click()
If textKriteria.Text = "" Then
    MsgBox "Silahkan isi data yang akan dicari!", vbExclamation + vbOKOnly, ""
    textKriteria.SetFocus
Else
    cmBatal.Caption = "&Tutup"
    FormPeminjamanBuku.AdodcUtama.Refresh
    With FormPeminjamanBuku.AdodcUtama.Recordset
        Select Case cmbCariDataBerdasarkan.ListIndex
        Case Is = 0
            .Find "NIA = '" & textKriteria.Text & "'"
        Case Is = 1
            .Find "NamaPeminjam = '" & textKriteria.Text & "'"
        Case Is = 2
            .Find "AlamatPeminjam = '" & textKriteria.Text & "'"
        Case Is = 3
            .Find "NoTelpPeminjam = '" & textKriteria.Text & "'"
        Case Is = 4
            .Find "StatusPendidikan = '" & textKriteria.Text & "'"
        Case Is = 5
            .Find "JudulBuku = '" & textKriteria.Text & "'"
        Case Is = 6
            .Find "KodeBuku = '" & textKriteria.Text & "'"
        Case Is = 7
            .Find "Kategori = '" & textKriteria.Text & "'"
        Case Is = 8
            .Find "Pengarang = '" & textKriteria.Text & "'"
        Case Is = 9
            .Find "Penerbit = '" & textKriteria.Text & "'"
        Case Is = 10
            .Find "TahunTerbit = '" & textKriteria.Text & "'"
        Case Is = 11
            .Find "Cetakan_Ke = '" & textKriteria.Text & "'"
        Case Is = 12
            .Find "HariPinjam = '" & textKriteria.Text & "'"
        Case Is = 13
            .Find "TanggalPinjam = '" & textKriteria.Text & "'"
        Case Is = 14
            .Find "BulanPinjam = '" & textKriteria.Text & "'"
        Case Is = 15
            .Find "TahunPinjam = '" & textKriteria.Text & "'"
        Case Is = 16
            .Find "Jumlah = '" & textKriteria.Text & "'"
        Case Is = 17
            .Find "LamaPinjam = '" & textKriteria.Text & "'"
        Case Is = 18
            .Find "SatuanTempo = '" & textKriteria.Text & "'"
        Case Is = 19
            .Find "Keterangan = '" & textKriteria.Text & "'"
        Case Is = 20
            .Find "NamaAdmin = '" & textKriteria.Text & "'"
        Case Is = 21
            .Find "Bagian = '" & textKriteria.Text & "'"
        Case Is = 22
            .Find "JamPinjam = '" & textKriteria.Text & "'"
        Case Is = 23
            .Find "MenitPinjam = '" & textKriteria.Text & "'"
        Case Is = 24
            .Find "DetikPinjam = '" & textKriteria.Text & "'"
        Case Is = 25
            .Find "SatuanWaktu = '" & textKriteria.Text & "'"
        Case Is = 26
            .Find "HariInputData = '" & textKriteria.Text & "'"
        End Select
        If .EOF Then
            MsgBox ("Data tidak ditemukan"), vbExclamation + vbOKOnly, ""
        Else
            Set FormPeminjamanBuku.DataGrid1.DataSource = FormPeminjamanBuku.AdodcUtama.Recordset
            cmBatal.Caption = "&Tutup"
            If FormPengaturan.cekTutupFormCari.Value = Checked Then Me.Hide
        End If
    End With
End If
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub

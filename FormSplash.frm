VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FormSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3120
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5430
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmTutup 
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2640
      Width           =   1095
   End
   Begin TabDlg.SSTab TabSplash 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4471
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "File"
      TabPicture(0)   =   "FormSplash.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FrameAnggota"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Transaksi"
      TabPicture(1)   =   "FormSplash.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FrameTransaksi"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Laporan"
      TabPicture(2)   =   "FormSplash.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameLaporan"
      Tab(2).ControlCount=   1
      Begin VB.Frame FrameLaporan 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   10
         Top             =   360
         Width           =   5175
         Begin VB.CommandButton cmLaporanDaftarAnggota 
            Caption         =   "Da&ftar Anggota"
            Height          =   735
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   1080
            Width           =   1575
         End
         Begin VB.CommandButton cmLaporanPeminjamanBuku 
            Caption         =   "Pe&minjaman Buku"
            Height          =   735
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmLaporanPengembalianBuku 
            Caption         =   "Pengemb&alian Buku"
            Height          =   735
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmLaporanBarangKoleksi 
            Caption         =   "Ba&rang Koleksi"
            Height          =   735
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame FrameTransaksi 
         Height          =   2055
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   5175
         Begin VB.CommandButton cmPengembalianBuku 
            Caption         =   "Pemgembalian Buku"
            Height          =   735
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmPeminjamanBuku 
            Caption         =   "&Peminjaman Buku"
            Height          =   735
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmBukuMasuk 
            Caption         =   "&Buku Masuk"
            Height          =   735
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmBarangKoleksiMasuk 
            Caption         =   "B&arang Koleksi"
            Height          =   735
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   1080
            Width           =   1575
         End
      End
      Begin VB.Frame FrameAnggota 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   1
         Top             =   360
         Width           =   5175
         Begin VB.CommandButton cmAnggota 
            Caption         =   "&Anggota"
            Height          =   735
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmDataBuku 
            Caption         =   "&Data Buku"
            Height          =   735
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Barang Koleksi"
            Height          =   735
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   240
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "FormSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmAnggota_Click()
    Unload Me
    With FormAnggota
        .Show
        .SetFocus
    End With
End Sub

Private Sub cmBarangKoleksiMasuk_Click()
    Unload Me
    With FormBarangKoleksiMasuk
        .Show
        .SetFocus
    End With
End Sub

Private Sub cmBukuMasuk_Click()
    Unload Me
    With FormBukuMasuk
        .Show
        .SetFocus
    End With
End Sub

Private Sub cmDataBuku_Click()
    Unload Me
    With FormManageBukuMasuk
        .Show
        .SetFocus
    End With
End Sub

Private Sub cmLaporanBarangKoleksi_Click()
    Unload Me
    With FormLaporanBarangKoleksi
        .Show
        .SetFocus
    End With
End Sub

Private Sub cmLaporanDaftarAnggota_Click()
    Unload Me
    With FormLaporanDaftarAnggota
        .Show
        .SetFocus
    End With
End Sub

Private Sub cmLaporanPeminjamanBuku_Click()
    Unload Me
    With FormLaporanPeminjamanBuku
        .Show
        .SetFocus
    End With
End Sub

Private Sub cmLaporanPengembalianBuku_Click()
    Unload Me
    With FormLaporanPengembalianBuku
        .Show
        .SetFocus
    End With
End Sub

Private Sub cmPeminjamanBuku_Click()
    Unload Me
    With FormPeminjamanBuku
        .Show
        .SetFocus
    End With
End Sub

Private Sub cmPengembalianBuku_Click()
    Unload Me
    With FormPengembalianBuku
        .Show
        .SetFocus
    End With
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Unload Me
    With FormManageBarangKoleksiMasuk
        .Show
        .SetFocus
    End With
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
    If FORM_UTAMA.StatusBawah.Panels.Item(2).Text = "Anonymous" Then
        With TabSplash
            .TabVisible(1) = False
            .TabVisible(2) = False
        End With
    End If
End Sub

VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormDaftarPenggunaBaru 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daftar Pengguna Baru"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6960
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormDaftarPenggunaBaru.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmLihatPengguna 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Lihat"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmTutup 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Batal"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmSimpan 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Simpan"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7560
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc AdodcTBLogin 
      Height          =   330
      Left            =   5760
      Top             =   0
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
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Data Identitas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   6735
      Begin VB.ComboBox cmbKategori 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox cmbStatusHubungan 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   4320
         Width           =   2895
      End
      Begin VB.ComboBox cmbStatusPekerjaan 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   3840
         Width           =   2895
      End
      Begin VB.ComboBox cmbStatusPendidikan 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   3360
         Width           =   2895
      End
      Begin VB.TextBox textTahunLahir 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5640
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   1800
         Width           =   950
      End
      Begin VB.ComboBox cmbBulanLahir 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1800
         Width           =   1215
      End
      Begin VB.ComboBox cmbTanggalLahir 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1800
         Width           =   735
      End
      Begin VB.ComboBox cmbJenisKelamin 
         Height          =   360
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1320
         Width           =   3855
      End
      Begin VB.TextBox textAlamat 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   2760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Text            =   "FormDaftarPenggunaBaru.frx":000C
         Top             =   2280
         Width           =   3855
      End
      Begin VB.TextBox textTempatLahir 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2760
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1800
         Width           =   950
      End
      Begin VB.TextBox textNamaAsli 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2760
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   840
         Width           =   3855
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kategori"
         Height          =   240
         Left            =   240
         TabIndex        =   40
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Hubungan"
         Height          =   240
         Left            =   240
         TabIndex        =   24
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   240
         Left            =   2640
         TabIndex        =   23
         Top             =   3840
         Width           =   45
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Pekerjaan"
         Height          =   240
         Left            =   240
         TabIndex        =   22
         Top             =   3840
         Width           =   1185
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   240
         Left            =   2640
         TabIndex        =   21
         Top             =   3360
         Width           =   45
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Pendidikan"
         Height          =   240
         Left            =   240
         TabIndex        =   20
         Top             =   3360
         Width           =   1245
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   240
         Left            =   2640
         TabIndex        =   19
         Top             =   2880
         Width           =   45
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         Height          =   240
         Left            =   240
         TabIndex        =   18
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   240
         Left            =   2640
         TabIndex        =   17
         Top             =   1800
         Width           =   45
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Lahir"
         Height          =   240
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   990
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   240
         Left            =   2640
         TabIndex        =   14
         Top             =   1320
         Width           =   45
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Asli"
         Height          =   240
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   705
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   240
         Left            =   2640
         TabIndex        =   11
         Top             =   360
         Width           =   45
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Kelamin"
         Height          =   240
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   960
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   240
         Left            =   2640
         TabIndex        =   9
         Top             =   840
         Width           =   45
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Data Login"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.TextBox TextPasswordLama 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   2760
         PasswordChar    =   "*"
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   840
         Width           =   3855
      End
      Begin VB.CheckBox CekBintang 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Char. Bintang"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   38
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.TextBox TextKonfirmasiPassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   2760
         PasswordChar    =   "*"
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox textPasswordBaru 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   2760
         PasswordChar    =   "*"
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1320
         Width           =   3855
      End
      Begin VB.TextBox textNamaPenggunaBaru 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2760
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password Lama"
         Height          =   240
         Left            =   240
         TabIndex        =   44
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   240
         Left            =   2640
         TabIndex        =   43
         Top             =   840
         Width           =   45
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   240
         Left            =   2640
         TabIndex        =   41
         Top             =   1800
         Width           =   45
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   240
         Left            =   2640
         TabIndex        =   26
         Top             =   1320
         Width           =   45
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Konfirmasi Password"
         Height          =   240
         Left            =   240
         TabIndex        =   25
         Top             =   1800
         Width           =   1470
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   240
         Left            =   2640
         TabIndex        =   5
         Top             =   840
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password Baru"
         Height          =   240
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   240
         Left            =   2640
         TabIndex        =   2
         Top             =   360
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pengguna Baru"
         Height          =   240
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1590
      End
   End
End
Attribute VB_Name = "FormDaftarPenggunaBaru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub SambungkanAdodc()
    NyambunggUtama
    With AdodcTBLogin
        .ConnectionString = CN.ConnectionString
        .RecordSource = "select * From tbLogin"
        .Refresh
    End With
End Sub
Sub AturKontrol()
    For Each Objek In Me
        If TypeName(Objek) = "TextBox" Then
            With Objek
                .Text = ""
                .MaxLength = 254
            End With
        End If
    Next
    With Me
        .cmbTanggalLahir.Clear
        .cmbBulanLahir.Clear
        .cmbJenisKelamin.Clear
        .cmbStatusPendidikan.Clear
        .cmbStatusPekerjaan.Clear
        .cmbStatusHubungan.Clear
            .cmbTanggalLahir.AddItem "01"
            .cmbTanggalLahir.AddItem "02"
            .cmbTanggalLahir.AddItem "03"
            .cmbTanggalLahir.AddItem "04"
            .cmbTanggalLahir.AddItem "05"
            .cmbTanggalLahir.AddItem "06"
            .cmbTanggalLahir.AddItem "07"
            .cmbTanggalLahir.AddItem "08"
            .cmbTanggalLahir.AddItem "09"
            .cmbTanggalLahir.AddItem "10"
            .cmbTanggalLahir.AddItem "11"
            .cmbTanggalLahir.AddItem "12"
            .cmbTanggalLahir.AddItem "13"
            .cmbTanggalLahir.AddItem "14"
            .cmbTanggalLahir.AddItem "15"
            .cmbTanggalLahir.AddItem "16"
            .cmbTanggalLahir.AddItem "17"
            .cmbTanggalLahir.AddItem "18"
            .cmbTanggalLahir.AddItem "19"
            .cmbTanggalLahir.AddItem "20"
            .cmbTanggalLahir.AddItem "21"
            .cmbTanggalLahir.AddItem "22"
            .cmbTanggalLahir.AddItem "23"
            .cmbTanggalLahir.AddItem "24"
            .cmbTanggalLahir.AddItem "25"
            .cmbTanggalLahir.AddItem "26"
            .cmbTanggalLahir.AddItem "27"
            .cmbTanggalLahir.AddItem "28"
            .cmbTanggalLahir.AddItem "29"
            .cmbTanggalLahir.AddItem "30"
            .cmbTanggalLahir.AddItem "31"
            .cmbTanggalLahir.Text = "01"
            .cmbBulanLahir.AddItem "Januari", 0
            .cmbBulanLahir.AddItem "Februari", 1
            .cmbBulanLahir.AddItem "Maret", 2
            .cmbBulanLahir.AddItem "April", 3
            .cmbBulanLahir.AddItem "Mei", 4
            .cmbBulanLahir.AddItem "Juni", 5
            .cmbBulanLahir.AddItem "Juli", 6
            .cmbBulanLahir.AddItem "Agustus", 7
            .cmbBulanLahir.AddItem "September", 8
            .cmbBulanLahir.AddItem "Oktober", 9
            .cmbBulanLahir.AddItem "November", 10
            .cmbBulanLahir.AddItem "Desember", 11
            .cmbBulanLahir.ListIndex = 1
            .cmbJenisKelamin.AddItem "Pria", 0
            .cmbJenisKelamin.AddItem "Wanita", 1
            .cmbJenisKelamin.ListIndex = 0
            .cmbStatusPendidikan.AddItem "SD", 0
            .cmbStatusPendidikan.AddItem "SMP", 1
            .cmbStatusPendidikan.AddItem "SMA", 2
            .cmbStatusPendidikan.AddItem "D1", 3
            .cmbStatusPendidikan.AddItem "D3", 4
            .cmbStatusPendidikan.AddItem "S1", 5
            .cmbStatusPendidikan.AddItem "S2", 6
            .cmbStatusPendidikan.AddItem "S3", 7
            .cmbStatusPendidikan.ListIndex = 4
            .cmbStatusPekerjaan.AddItem "Pelajar/Mahasiswa", 0
            .cmbStatusPekerjaan.AddItem "Bekerja", 1
            .cmbStatusPekerjaan.ListIndex = 0
            .cmbStatusHubungan.AddItem "Lajang", 0
            .cmbStatusHubungan.AddItem "Menikah", 1
            .cmbStatusHubungan.ListIndex = 0
    End With
        If AdodcTBLogin.Recordset.RecordCount = 0 Then
            cmLihatPengguna.Enabled = False
        Else
            cmLihatPengguna.Enabled = True
        End If
        With Me
            .cmbKategori.Clear
            .cmbKategori.AddItem "Admin", 0
            .cmbKategori.AddItem "Pegawai/Staff", 1
            .cmbKategori.ListIndex = 0
        End With
        If FORM_UTAMA.StatusBawah.Panels.Item(2).Text = "" Then
        Else
            Unload FormAutorisasi
        End If
End Sub
Sub SimpanDataPengguna()
On Error GoTo HancurkanError
    If cmSimpan.Caption = "&Simpan" Then
        With AdodcTBLogin
            .Recordset.AddNew
            .Recordset.Fields(0).Value = textNamaPenggunaBaru.Text
            .Recordset.Fields(1).Value = textPasswordBaru.Text
            .Recordset.Fields(2).Value = cmbKategori.Text
            .Recordset.Fields(3).Value = textNamaAsli.Text
            .Recordset.Fields(4).Value = cmbJenisKelamin.Text
            .Recordset.Fields(5).Value = textTempatLahir.Text
            .Recordset.Fields(6).Value = cmbTanggalLahir.Text
            .Recordset.Fields(7).Value = cmbBulanLahir.Text
            .Recordset.Fields(8).Value = textTahunLahir.Text
            .Recordset.Fields(9).Value = textAlamat.Text
            .Recordset.Fields(10).Value = cmbStatusPendidikan.Text
            .Recordset.Fields(11).Value = cmbStatusPekerjaan.Text
            .Recordset.Fields(12).Value = cmbStatusHubungan.Text
            .Recordset.Update
            .Refresh
        End With
        MsgBox "Data Pengguna dengan Nama : '" & textNamaPenggunaBaru.Text & "/" & textNamaAsli.Text & "' berhasil didaftarkan!", vbInformation + vbOKOnly, "Data Sukses Disimpan!"
        cmTutup.Caption = "&Tutup"
        AturKontrol
        If FORM_UTAMA.StatusBawah.Panels.Item(2).Text = "" Then
        Else
            FormDaftarAdmin.AturKontrol
            Unload Me
        End If
    ElseIf cmSimpan.Caption = "&Perbarui" Then
        If TextPasswordLama.Text <> FormDaftarAdmin.AdodcUtamaPWTampilkan.Recordset.Fields(1).Value Then
            MsgBox "Password lama Anda tidak sesuai!", vbCritical + vbOKOnly, "Error"
            TextPasswordLama.SetFocus
        Else
            With FormDaftarAdmin.AdodcUtamaPWTampilkan
                .Recordset.Delete
                .Recordset.AddNew
                .Recordset.Fields(0).Value = textNamaPenggunaBaru.Text
                .Recordset.Fields(1).Value = textPasswordBaru.Text
                .Recordset.Fields(2).Value = cmbKategori.Text
                .Recordset.Fields(3).Value = textNamaAsli.Text
                .Recordset.Fields(4).Value = cmbJenisKelamin.Text
                .Recordset.Fields(5).Value = textTempatLahir.Text
                .Recordset.Fields(6).Value = cmbTanggalLahir.Text
                .Recordset.Fields(7).Value = cmbBulanLahir.Text
                .Recordset.Fields(8).Value = textTahunLahir.Text
                .Recordset.Fields(9).Value = textAlamat.Text
                .Recordset.Fields(10).Value = cmbStatusPendidikan.Text
                .Recordset.Fields(11).Value = cmbStatusPekerjaan.Text
                .Recordset.Fields(12).Value = cmbStatusHubungan.Text
                .Recordset.Update
                .Refresh
            End With
            MsgBox "Data Pengguna dengan Nama : '" & textNamaPenggunaBaru.Text & "/" & textNamaAsli.Text & "' berhasil diperbarui!", vbInformation + vbOKOnly, "Data Sukses Diperbarui!"
            cmTutup.Caption = "&Tutup"
            AturKontrol
            If FORM_UTAMA.StatusBawah.Panels.Item(2).Text = "" Then
            Else
                FormDaftarAdmin.AturKontrol
                Unload Me
            End If
        End If
    End If
Exit Sub
HancurkanError:
    PusatError
End Sub



Private Sub CekBintang_Click()
If TextPasswordLama.Enabled = False Then
    Select Case CekBintang.Value
        Case Is = Checked
            textPasswordBaru.PasswordChar = "*"
            TextKonfirmasiPassword.PasswordChar = "*"
        Case Is = Unchecked
            textPasswordBaru.PasswordChar = ""
            TextKonfirmasiPassword.PasswordChar = ""
        End Select
ElseIf TextPasswordLama.Enabled = True Then
    Select Case CekBintang.Value
        Case Is = Checked
            TextPasswordLama.PasswordChar = "*"
            textPasswordBaru.PasswordChar = "*"
            TextKonfirmasiPassword.PasswordChar = "*"
        Case Is = Unchecked
            TextPasswordLama.PasswordChar = ""
            textPasswordBaru.PasswordChar = ""
            TextKonfirmasiPassword.PasswordChar = ""
        End Select
End If
End Sub

Private Sub cmLihatPengguna_Click()
    FormLihatPengguna.Show vbModal, Me
End Sub

Private Sub cmSimpan_Click()
    If textNamaPenggunaBaru.Text = "" Then
        MsgBox "Silahkan Isi Nama Pengguna Baru yang akan digunakan untuk Login", vbExclamation + vbOKOnly, "MainSystem : Nama Pengguna Baru?"
        textNamaPenggunaBaru.SetFocus
    ElseIf textPasswordBaru.Text = "" Then
        MsgBox "Silahkan Isi Nama PASSWORD Pengguna Baru yang akan digunakan untuk Login", vbExclamation + vbOKOnly, "MainSystem : Password Baru?"
        textPasswordBaru.SetFocus
    ElseIf TextKonfirmasiPassword.Text = "" Then
        MsgBox "Silahkan KONFIRMASIKAN PASSWORD Pengguna Baru yang akan digunakan untuk Login", vbExclamation + vbOKOnly, "MainSystem : Konfirmasi Password?"
        TextKonfirmasiPassword.SetFocus
    ElseIf TextKonfirmasiPassword.Text <> textPasswordBaru.Text Then
        MsgBox "Maaf Password Gagal Di Konfirmasi dan Tidak Sesuai!", vbCritical + vbOKOnly, "MainSystem : Password Tidak Sesuai"
        TextKonfirmasiPassword.SetFocus
    ElseIf textNamaAsli.Text = "" Then
        MsgBox "Silahkan isi Nama Asli dari Pengguna", vbExclamation + vbOKOnly, "MainSystem"
        textNamaAsli.SetFocus
    ElseIf textTempatLahir.Text = "" Then
        MsgBox "Silahkan isi Tempat/Kota Lahir Pengguna", vbExclamation + vbOKOnly, "MainSystem"
        textTempatLahir.SetFocus
    ElseIf textTahunLahir.Text = "" Then
        MsgBox "Silahkan isi Tahun Lahir Pengguna", vbExclamation + vbOKOnly, "MainSystem"
        textTahunLahir.SetFocus
    ElseIf textAlamat.Text = "" Then
        MsgBox "Silahkan Alamat Pengguna", vbExclamation + vbOKOnly, "MainSystem"
        textAlamat.SetFocus
    Else
        If CekBintang.Value = Unchecked Then
            MsgBox "Maaf, data hanya bisa disimpan saat character PASSWORD disembunyikan!", vbCritical + vbOKOnly, "MainSystem"
            CekBintang.SetFocus
        Else
            X = MsgBox("Data-Data telah diverifikasi," & vbCrLf & _
                        "Anda yakin ingin menyimpan data pengguna ini?", vbQuestion + vbYesNo, "Konfirmasi")
            If X = vbYes Then SimpanDataPengguna
        End If
    End If
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    SambungkanAdodc
    AturKontrol
End Sub


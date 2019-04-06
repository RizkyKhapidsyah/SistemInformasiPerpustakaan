VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormAutorisasi 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autentikasi"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5025
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormAutorisasi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AdodcUtama 
      Height          =   330
      Left            =   120
      Top             =   2040
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
   Begin VB.TextBox textAutentikasiSembunyi 
      Height          =   390
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton cmBatal 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Batal"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmOK 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&OK"
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.TextBox textAutentikasi 
         Height          =   390
         Left            =   2280
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Masukkan Password Anda"
         Height          =   270
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1710
      End
   End
   Begin VB.Label LabelJenisAksi 
      Caption         =   "Label3"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   855
   End
End
Attribute VB_Name = "FormAutorisasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    NyambunggUtama
    With textAutentikasi
        .Text = ""
        .PasswordChar = "*"
    End With
    With AdodcUtama
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * From tbLogin"
        .Refresh
    End With
End Sub

Private Sub cmBatal_Click()
    Unload Me
End Sub

Private Sub cmOK_Click()
On Error GoTo HancurkanError
    If AdodcUtama.Recordset.State = 1 Then AdodcUtama.Recordset.Close
    AdodcUtama.Recordset.Open "select * from tbLogin where NamaPengguna= '" & textAutentikasiSembunyi.Text & "' And Password = '" & textAutentikasi.Text & "'", CN, 3, 3
        If textAutentikasi.Text = "" Then
            MsgBox "Silahkan Isi Password Anda!", vbExclamation + vbOKOnly, "MainSystem : Autentikasi"
            textAutentikasi.SetFocus
        Else
            If AdodcUtama.Recordset.EOF Then
                MsgBox "Maaf, Password Anda Salah!" & vbCrLf & _
                        "Silahkan coba lagi!", vbCritical + vbOKOnly, ""
                textAutentikasi.SetFocus
            Else
                Select Case LabelJenisAksi.Caption
                Case "TampilkanFormDaftarAdmin"
                    Unload Me
                    With FormDaftarAdmin
                        .Show
                        .SetFocus
                    End With
                Case "&Tambah"
                    With FormDaftarPenggunaBaru
                        .Caption = "Tambah Data - Admin"
                        .TextPasswordLama.Enabled = False
                        .TextPasswordLama.BackColor = FormDaftarPenggunaBaru.BackColor
                        .cmSimpan.Caption = "&Simpan"
                        .cmTutup.Caption = "&Batal"
                        .cmLihatPengguna.Enabled = False
                        .Show vbModal, Me
                    End With
                Case "&Edit"
                    With FormDaftarPenggunaBaru
                        .Caption = "Edit Data - Admin"
                        .TextPasswordLama.Enabled = True
                        .TextPasswordLama.BackColor = vbWhite
                        .cmSimpan.Caption = "&Perbarui"
                        .cmTutup.Caption = "&Batal"
                        .cmLihatPengguna.Enabled = False
                        
                        .textNamaPenggunaBaru.Text = FormDaftarAdmin.AdodcUtamaPWTampilkan.Recordset.Fields(0).Value
                        .TextPasswordLama.Text = ""
                        .textPasswordBaru.Text = ""
                        .TextKonfirmasiPassword.Text = ""
                        .cmbKategori.Text = FormDaftarAdmin.AdodcUtamaPWTampilkan.Recordset.Fields(2).Value
                        .textNamaAsli.Text = FormDaftarAdmin.AdodcUtamaPWTampilkan.Recordset.Fields(3).Value
                        .cmbJenisKelamin.Text = FormDaftarAdmin.AdodcUtamaPWTampilkan.Recordset.Fields(4).Value
                        .textTempatLahir.Text = FormDaftarAdmin.AdodcUtamaPWTampilkan.Recordset.Fields(5).Value
                        .cmbTanggalLahir.Text = FormDaftarAdmin.AdodcUtamaPWTampilkan.Recordset.Fields(6).Value
                        .cmbBulanLahir.Text = FormDaftarAdmin.AdodcUtamaPWTampilkan.Recordset.Fields(7).Value
                        .textTahunLahir.Text = FormDaftarAdmin.AdodcUtamaPWTampilkan.Recordset.Fields(8).Value
                        .textAlamat.Text = FormDaftarAdmin.AdodcUtamaPWTampilkan.Recordset.Fields(9).Value
                        .cmbStatusPendidikan.Text = FormDaftarAdmin.AdodcUtamaPWTampilkan.Recordset.Fields(10).Value
                        .cmbStatusPekerjaan.Text = FormDaftarAdmin.AdodcUtamaPWTampilkan.Recordset.Fields(11).Value
                        .cmbStatusHubungan.Text = FormDaftarAdmin.AdodcUtamaPWTampilkan.Recordset.Fields(12).Value
                        
                        .Show vbModal, Me
                    End With
                Case "&Hapus"
                    X = MsgBox("Apakah Anda yakin ingin menghapus data Admin dengan nama : '" & FormDaftarAdmin.AdodcUtamaPWTampilkan.Recordset.Fields(0).Value & "' ?", vbQuestion + vbYesNo, "Hapus?")
                    If X = vbYes Then
                        If textAutentikasiSembunyi.Text = FORM_UTAMA.StatusBawah.Panels.Item(2).Text Then
                            MsgBox "Maaf, data tidak dapat dihapus. (Admin:'" & textAutentikasiSembunyi.Text & "'" & vbCrLf & _
                                    "Karena Anda (" & textAutentikasiSembunyi.Text & ") adalah admin yang aktif saat ini di program!", vbCritical + vbOKOnly, "Main System - Akses Gagal"
                        Else
                            With FormDaftarAdmin.AdodcUtamaPWTampilkan
                                .Recordset.Delete
                                .Refresh
                            End With
                            Unload Me
                        End If
                    End If
                End Select
            End If
        End If
Exit Sub
HancurkanError:
    PusatError
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub

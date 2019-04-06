VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormBukuMasuk 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buku Masuk"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6825
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormBukuMasuk.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   6825
   Begin MSAdodcLib.Adodc AdodcUtama 
      Height          =   330
      Left            =   4080
      Top             =   5640
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
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Reset"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmManage 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Manage"
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmSimpan 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmTutup 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Tutup"
      Height          =   495
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   120
      TabIndex        =   22
      Top             =   0
      Width           =   6615
      Begin VB.CommandButton cmTambahJenisPerusahaan 
         BackColor       =   &H00E0E0E0&
         Caption         =   "+"
         Height          =   375
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   720
         Width           =   495
      End
      Begin VB.ComboBox cmbJenis 
         Height          =   390
         ItemData        =   "FormBukuMasuk.frx":000C
         Left            =   4200
         List            =   "FormBukuMasuk.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   720
         Width           =   1755
      End
      Begin VB.CommandButton cmSet 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Set"
         Height          =   375
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmbTahun 
         Height          =   390
         ItemData        =   "FormBukuMasuk.frx":0010
         Left            =   4200
         List            =   "FormBukuMasuk.frx":0012
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   240
         Width           =   1275
      End
      Begin VB.TextBox textPengirim 
         DataField       =   "Kode"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         TabIndex        =   27
         Top             =   720
         Width           =   2235
      End
      Begin VB.ComboBox cmbHari 
         Height          =   390
         ItemData        =   "FormBukuMasuk.frx":0014
         Left            =   1440
         List            =   "FormBukuMasuk.frx":0016
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   240
         Width           =   1035
      End
      Begin VB.ComboBox cmbTanggal 
         Height          =   390
         ItemData        =   "FormBukuMasuk.frx":0018
         Left            =   2520
         List            =   "FormBukuMasuk.frx":001A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   240
         Width           =   795
      End
      Begin VB.ComboBox cmbBulan 
         Height          =   390
         ItemData        =   "FormBukuMasuk.frx":001C
         Left            =   3360
         List            =   "FormBukuMasuk.frx":001E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis"
         Height          =   270
         Left            =   3720
         TabIndex        =   31
         Top             =   780
         Width           =   345
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Pengirim"
         Height          =   270
         Left            =   120
         TabIndex        =   28
         Top             =   780
         Width           =   570
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Masuk"
         Height          =   270
         Left            =   120
         TabIndex        =   26
         Top             =   300
         Width           =   930
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   6615
      Begin VB.CommandButton cmAcakKode 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Acak"
         Height          =   390
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox textKode 
         DataField       =   "Kode"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         TabIndex        =   11
         Top             =   240
         Width           =   2115
      End
      Begin VB.ComboBox cmbKategori 
         Height          =   390
         ItemData        =   "FormBukuMasuk.frx":0020
         Left            =   1440
         List            =   "FormBukuMasuk.frx":0022
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1080
         Width           =   2115
      End
      Begin VB.TextBox textPengarang 
         DataField       =   "Pengarang"
         DataSource      =   "adcBuku"
         Height          =   790
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   1500
         Width           =   4995
      End
      Begin VB.TextBox textKeterangan 
         DataField       =   "Keterangan"
         DataSource      =   "adcBuku"
         Height          =   735
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   3630
         Width           =   4995
      End
      Begin VB.TextBox textStok 
         Alignment       =   2  'Center
         DataField       =   "Stok"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   4800
         TabIndex        =   7
         Top             =   3200
         Width           =   1635
      End
      Begin VB.TextBox textJumlah 
         Alignment       =   1  'Right Justify
         DataField       =   "Jumlah"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         TabIndex        =   6
         Top             =   3200
         Width           =   2085
      End
      Begin VB.TextBox textCetakanKe 
         Alignment       =   2  'Center
         DataField       =   "Cetakan"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   4800
         TabIndex        =   5
         Top             =   2760
         Width           =   1635
      End
      Begin VB.TextBox textTahunTerbit 
         Alignment       =   1  'Right Justify
         DataField       =   "TahunTerbit"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         TabIndex        =   4
         Top             =   2760
         Width           =   2085
      End
      Begin VB.TextBox textPenerbit 
         DataField       =   "Penerbit"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         TabIndex        =   3
         Top             =   2340
         Width           =   4995
      End
      Begin VB.TextBox textJudul 
         DataField       =   "Judul"
         DataSource      =   "adcBuku"
         Height          =   390
         Left            =   1440
         TabIndex        =   2
         Top             =   660
         Width           =   4995
      End
      Begin VB.CommandButton cmTambahKategori 
         BackColor       =   &H00E0E0E0&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Tambah Kategori Buku"
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cetakan ke"
         Height          =   270
         Left            =   3960
         TabIndex        =   21
         Top             =   2760
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
         Height          =   270
         Left            =   120
         TabIndex        =   20
         Top             =   300
         Width           =   315
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Kategori"
         Height          =   270
         Left            =   120
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   3630
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stok"
         Height          =   270
         Left            =   4440
         TabIndex        =   16
         Top             =   3200
         Width           =   285
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah"
         Height          =   270
         Left            =   120
         TabIndex        =   15
         Top             =   3200
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tahun Terbit"
         Height          =   270
         Left            =   120
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   2400
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Judul"
         Height          =   270
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   345
      End
   End
   Begin MSAdodcLib.Adodc AdodcUntukCMBJenis 
      Height          =   330
      Left            =   4080
      Top             =   5880
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
   Begin MSAdodcLib.Adodc AdodcUntukCMBKategori 
      Height          =   330
      Left            =   4080
      Top             =   6120
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
Attribute VB_Name = "FormBukuMasuk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    NyambunggUtama
    With AdodcUtama
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * From tbBuku order by JudulBuku asc;"
        .Refresh
    End With
    With AdodcUntukCMBJenis
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * From tbJenisPerusahaan "
        .Refresh
    End With
    With AdodcUntukCMBKategori
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * From tbKategori "
        .Refresh
    End With
    With cmbHari
        .Clear
        .AddItem "Senin", 0
        .AddItem "Selasa", 1
        .AddItem "Rabu", 2
        .AddItem "Kamis", 3
        .AddItem "Jumat", 4
        .AddItem "Sabtu", 5
        .AddItem "Minggu", 6
        .ListIndex = 0
    End With
    With cmbBulan
        .Clear
        .AddItem "1", 0
        .AddItem "2", 1
        .AddItem "3", 2
        .AddItem "4", 3
        .AddItem "5", 4
        .AddItem "6", 5
        .AddItem "7", 6
        .AddItem "8", 7
        .AddItem "9", 8
        .AddItem "10", 9
        .AddItem "11", 10
        .AddItem "12", 11
        .ListIndex = 0
    End With
    cmbTahun.Clear
    For X = 1800 To 3000
        cmbTahun.AddItem X
    Next
    cmbTahun.Text = Year(Date)
    MasukkanDatabaseKeAdodcUntukCMBJenis
    MasukkanDatabaseKeAdodcUntukCMBKategori
    ResetInput
End Sub
Sub MasukkanDatabaseKeAdodcUntukCMBJenis()
    AdodcUntukCMBJenis.Refresh
    cmbJenis.Clear
        Do Until AdodcUntukCMBJenis.Recordset.EOF
            cmbJenis.AddItem AdodcUntukCMBJenis.Recordset.Fields(0).Value, 0
            AdodcUntukCMBJenis.Recordset.MoveNext
        Loop
        AdodcUntukCMBJenis.Refresh
        cmbJenis.ListIndex = 0
End Sub
Sub MasukkanDatabaseKeAdodcUntukCMBKategori()
    AdodcUntukCMBKategori.Refresh
    cmbKategori.Clear
        Do Until AdodcUntukCMBKategori.Recordset.EOF
            cmbKategori.AddItem AdodcUntukCMBKategori.Recordset.Fields(1).Value, 0
            AdodcUntukCMBKategori.Recordset.MoveNext
        Loop
        AdodcUntukCMBKategori.Refresh
        cmbKategori.ListIndex = 0
End Sub
Sub ResetInput()
For Each Objek In Me
    If TypeName(Objek) = "TextBox" Then
        With Objek
            .MaxLength = 254
            .Text = ""
        End With
    End If
Next
textKode.Text = Val(AdodcUtama.Recordset.RecordCount) + 1 & "BPAD" & Second(Time) & Hour(Time) & Minute(Time) & Val(Second(Time) * 3)
End Sub


Private Sub cmAcakKode_Click()
    textKode.Text = Val(AdodcUtama.Recordset.RecordCount) + 1 & "BPAD" & Second(Time) & Hour(Time) & Minute(Time) & Val(Second(Time) * 3)
    textJudul.SetFocus
End Sub

Private Sub cmbBulan_Click()
    cmbTanggal.Clear
    Select Case cmbBulan.ListIndex
        Case Is = 0, 2, 4, 6, 7, 9, 11
            For X = 1 To 31
                With cmbTanggal
                    .AddItem X, 0
                End With
            Next
        Case Is = 1
            If Val(cmbTahun.Text) Mod 4 Then
                For X = 1 To 29
                    With cmbTanggal
                        .AddItem X, 0
                    End With
                Next
            Else
                For X = 1 To 28
                    With cmbTanggal
                        .AddItem X, 0
                    End With
                Next
            End If
        Case Is = 3, 5, 8, 10
            For X = 1 To 30
                With cmbTanggal
                    .AddItem X, 0
                End With
            Next
    End Select
    cmbTanggal.Text = "1"
End Sub

Private Sub cmManage_Click()
    With FormManageBukuMasuk
        .Show
        .SetFocus
    End With
End Sub

Private Sub cmReset_Click()
    AturKontrol
    ResetInput
End Sub

Private Sub cmSet_Click()
    cmbTanggal.Text = Day(Date)
    cmbBulan.Text = Month(Date)
    cmbTahun.Text = Year(Date)
    textPengirim.SetFocus
End Sub

Private Sub cmSimpan_Click()
    If textPengirim.Text = "" Then
        MsgBox "Silahkan isi nama pengirim buku", vbExclamation + vbOKOnly, ""
        textPengirim.SetFocus
    ElseIf textKode.Text = "" Then
        MsgBox "Silahkan isi Kode Buku atau klik 'Acak Kode'", vbExclamation + vbOKOnly, ""
        textKode.SetFocus
    ElseIf textJudul.Text = "" Then
        MsgBox "Silahkan isi judul buku yang akan diinput", vbExclamation + vbOKOnly, ""
        textJudul.SetFocus
    ElseIf textPengarang.Text = "" Then
        MsgBox "Silahkan isi nama penulis/pengarang buku", vbExclamation + vbOKOnly, ""
        textPengarang.SetFocus
    ElseIf textPenerbit.Text = "" Then
        MsgBox "Silahkan isi nama penerbit buku", vbExclamation + vbOKOnly, ""
        textPenerbit.SetFocus
    ElseIf textTahunTerbit.Text = "" Then
        MsgBox "Silahkan isi tahun terbit bukut", vbExclamation + vbOKOnly, ""
        textTahunTerbit.SetFocus
    ElseIf textCetakanKe.Text = "" Then
        MsgBox "Pada tahun berapakah buku dicetak?", vbExclamation + vbOKOnly, ""
        textCetakanKe.SetFocus
    ElseIf textJumlah.Text = "" Then
        MsgBox "Berapa jumlah buku yang dikirim?", vbExclamation + vbOKOnly, ""
        textJumlah.SetFocus
    ElseIf textStok.Text = "" Then
        MsgBox "Berapa jumlah stok buku yang dikirim?", vbExclamation + vbOKOnly, ""
        textStok.SetFocus
    ElseIf textKeterangan.Text = "" Then
        MsgBox "Silahkan isi keterangan yang dibutuhkan.", vbExclamation + vbOKOnly, ""
        textKeterangan.SetFocus
    Else
        Select Case cmSimpan.Caption
        Case "&Simpan"
            X = MsgBox("Anda yakin ingin menambahkan data buku dengan kode '" & textKode.Text & "' ?", vbQuestion + vbYesNo, "Konfirmasi")
            If X = vbYes Then
                With AdodcUtama
                    .Recordset.AddNew
                    .Recordset.Fields(0).Value = cmbHari.Text
                    .Recordset.Fields(1).Value = cmbTanggal.Text
                    .Recordset.Fields(2).Value = cmbBulan.Text
                    .Recordset.Fields(3).Value = cmbTahun.Text
                    .Recordset.Fields(4).Value = textPengirim.Text
                    .Recordset.Fields(5).Value = cmbJenis.Text
                    .Recordset.Fields(6).Value = textKode.Text
                    .Recordset.Fields(7).Value = textJudul.Text
                    .Recordset.Fields(8).Value = cmbKategori.Text
                    .Recordset.Fields(9).Value = textPengarang.Text
                    .Recordset.Fields(10).Value = textPenerbit.Text
                    .Recordset.Fields(11).Value = textTahunTerbit.Text
                    .Recordset.Fields(12).Value = textCetakanKe.Text
                    .Recordset.Fields(13).Value = textJumlah.Text
                    .Recordset.Fields(14).Value = textStok.Text
                    .Recordset.Fields(15).Value = textKeterangan.Text
                    .Recordset.Update
                    .Refresh
                End With
                AturKontrol
                ResetInput
                cmTutup.Caption = "&Tutup"
                textPengirim.SetFocus
                FormManageBukuMasuk.AturKontrol
            End If
        Case "&Perbarui"
            X = MsgBox("Anda yakin ingin memperbarui data buku dengan kode '" & textKode.Text & "' ?", vbQuestion + vbYesNo, "Konfirmasi")
            If X = vbYes Then
                With FormManageBukuMasuk.AdodcUtama
                    .Recordset.Delete
                    .Recordset.AddNew
                    .Recordset.Fields(0).Value = cmbHari.Text
                    .Recordset.Fields(1).Value = cmbTanggal.Text
                    .Recordset.Fields(2).Value = cmbBulan.Text
                    .Recordset.Fields(3).Value = cmbTahun.Text
                    .Recordset.Fields(4).Value = textPengirim.Text
                    .Recordset.Fields(5).Value = cmbJenis.Text
                    .Recordset.Fields(6).Value = textKode.Text
                    .Recordset.Fields(7).Value = textJudul.Text
                    .Recordset.Fields(8).Value = cmbKategori.Text
                    .Recordset.Fields(9).Value = textPengarang.Text
                    .Recordset.Fields(10).Value = textPenerbit.Text
                    .Recordset.Fields(11).Value = textTahunTerbit.Text
                    .Recordset.Fields(12).Value = textCetakanKe.Text
                    .Recordset.Fields(13).Value = textJumlah.Text
                    .Recordset.Fields(14).Value = textStok.Text
                    .Recordset.Fields(15).Value = textKeterangan.Text
                    .Recordset.Update
                    .Refresh
                End With
                AturKontrol
                ResetInput
                cmTutup.Caption = "&Tutup"
                cmSimpan.Caption = "&Simpan"
                textPengirim.SetFocus
                FormManageBukuMasuk.AturKontrol
            End If
        End Select
    End If
End Sub

Private Sub cmTambahJenisPerusahaan_Click()
With FormTambahJenisPerusahaan
    .Show
    .SetFocus
End With
End Sub

Private Sub cmTambahKategori_Click()
With formTambahKategoriBuku
    .Show
    .SetFocus
End With
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub

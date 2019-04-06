VERSION 5.00
Begin VB.Form FormFilterBarangKoleksiMasuk 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filter Data - Barang Koleksi Masuk"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6240
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormFilterBarangKoleksiMasuk.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   6240
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
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   6015
      Begin VB.ComboBox cmbFilterBerdasarkan 
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   3975
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   2535
      End
      Begin VB.CommandButton cmFilter 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Filter"
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
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filter Berdasarkan :"
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
         Top             =   360
         Width           =   1290
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dengan Mode :"
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
         Top             =   840
         Width           =   915
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
      Top             =   1440
      Width           =   1335
   End
End
Attribute VB_Name = "FormFilterBarangKoleksiMasuk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
With Me
    .cmbFilterBerdasarkan.Clear
    .cmbFilterBerdasarkan.AddItem FormManageBarangKoleksiMasuk.AdodcUtama.Recordset.Fields(0).Name, 0
    .cmbFilterBerdasarkan.AddItem FormManageBarangKoleksiMasuk.AdodcUtama.Recordset.Fields(1).Name, 1
    .cmbFilterBerdasarkan.AddItem FormManageBarangKoleksiMasuk.AdodcUtama.Recordset.Fields(2).Name, 2
    .cmbFilterBerdasarkan.AddItem FormManageBarangKoleksiMasuk.AdodcUtama.Recordset.Fields(3).Name, 3
    .cmbFilterBerdasarkan.AddItem FormManageBarangKoleksiMasuk.AdodcUtama.Recordset.Fields(4).Name, 4
    .cmbFilterBerdasarkan.AddItem FormManageBarangKoleksiMasuk.AdodcUtama.Recordset.Fields(5).Name, 5
    .cmbFilterBerdasarkan.AddItem FormManageBarangKoleksiMasuk.AdodcUtama.Recordset.Fields(6).Name, 6
    .cmbFilterBerdasarkan.AddItem FormManageBarangKoleksiMasuk.AdodcUtama.Recordset.Fields(7).Name, 7
    .cmbFilterBerdasarkan.AddItem FormManageBarangKoleksiMasuk.AdodcUtama.Recordset.Fields(8).Name, 8
    .cmbFilterBerdasarkan.ListIndex = 4
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
    With FormManageBarangKoleksiMasuk
        .AdodcUtama.Refresh
            Select Case cmbFilterBerdasarkan.ListIndex
            Case Is = 0
                .AdodcUtama.RecordSource = "Select Hari from tbBarangKoleksiMasuk order by Hari asc;"
            Case Is = 1
                .AdodcUtama.RecordSource = "Select Tanggal from tbBarangKoleksiMasuk order by Tanggal asc;"
            Case Is = 2
                .AdodcUtama.RecordSource = "Select Bulan from tbBarangKoleksiMasuk order by Bulan asc;"
            Case Is = 3
                .AdodcUtama.RecordSource = "Select Tahun from tbBarangKoleksiMasuk order by Tahun asc;"
            Case Is = 4
                .AdodcUtama.RecordSource = "Select Kode_Barang from tbBarangKoleksiMasuk order by Kode_Barang asc;"
            Case Is = 5
                .AdodcUtama.RecordSource = "Select Nama_Barang from tbBarangKoleksiMasuk order by Nama_Barang asc;"
            Case Is = 6
                .AdodcUtama.RecordSource = "Select Status_Barang from tbBarangKoleksiMasuk order by Status_Barang asc;"
            Case Is = 7
                .AdodcUtama.RecordSource = "Select Penerima_(Nama_Admin) from tbBarangKoleksiMasuk order by Penerima_(Nama_Admin) asc;"
            Case Is = 8
                .AdodcUtama.RecordSource = "Select Keterangan from tbBarangKoleksiMasuk order by Keterangan asc;"
            End Select
    End With
ElseIf cmbMode.ListIndex = 1 Then
    With FormManageBarangKoleksiMasuk
        .AdodcUtama.Refresh
            Select Case cmbFilterBerdasarkan.ListIndex
            Case Is = 0
                .AdodcUtama.RecordSource = "Select Hari from tbBarangKoleksiMasuk order by Hari desc;"
            Case Is = 1
                .AdodcUtama.RecordSource = "Select Tanggal from tbBarangKoleksiMasuk order by Tanggal desc;"
            Case Is = 2
                .AdodcUtama.RecordSource = "Select Bulan from tbBarangKoleksiMasuk order by Bulan desc;"
            Case Is = 3
                .AdodcUtama.RecordSource = "Select Tahun from tbBarangKoleksiMasuk order by Tahun desc;"
            Case Is = 4
                .AdodcUtama.RecordSource = "Select Kode_Barang from tbBarangKoleksiMasuk order by Kode_Barang desc;"
            Case Is = 5
                .AdodcUtama.RecordSource = "Select Nama_Barang from tbBarangKoleksiMasuk order by Nama_Barang desc;"
            Case Is = 6
                .AdodcUtama.RecordSource = "Select Status_Barang from tbBarangKoleksiMasuk order by Status_Barang desc;"
            Case Is = 7
                .AdodcUtama.RecordSource = "Select Penerima_(Nama_Admin) from tbBarangKoleksiMasuk order by Penerima_(Nama_Admin) desc;"
            Case Is = 8
                .AdodcUtama.RecordSource = "Select Keterangan from tbBarangKoleksiMasuk order by Keterangan desc;"
            End Select
    End With
End If
    FormManageBarangKoleksiMasuk.AdodcUtama.Refresh
    cmBatal.Caption = "&Tutup"
    If FormPengaturan.cekTutupFormFilter.Value = Checked Then Me.Hide
    With FormManageBarangKoleksiMasuk
        .cmEdit.Enabled = False
        .cmCari.Enabled = False
        .cmSorot.Enabled = False
        .cmFilter.Enabled = False
        .cmHapus.Enabled = False
        .cmTambah.Enabled = False
    End With
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub



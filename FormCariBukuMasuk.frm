VERSION 5.00
Begin VB.Form FormCariBukuMasuk 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cari Data - Buku Masuk"
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
   Icon            =   "FormCariBukuMasuk.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   6000
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
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
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
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.CommandButton cmCari 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cari"
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
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox textKriteria 
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
         Left            =   2040
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   720
         Width           =   2415
      End
      Begin VB.ComboBox cmbCariDataBerdasarkan 
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
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dengan Kriteria"
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
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cari data berdasarkan"
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
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FormCariBukuMasuk"
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
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(0).Name, 0
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(1).Name, 1
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(2).Name, 2
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(3).Name, 3
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(4).Name, 4
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(5).Name, 5
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(6).Name, 6
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(7).Name, 7
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(8).Name, 8
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(9).Name, 9
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(10).Name, 10
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(11).Name, 11
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(12).Name, 12
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(13).Name, 13
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(14).Name, 14
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(15).Name, 15
        .ListIndex = 6
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
    FormManageBukuMasuk.AdodcUtama.Refresh
    With FormManageBukuMasuk.AdodcUtama.Recordset
        Select Case cmbCariDataBerdasarkan.ListIndex
        Case Is = 0
            .Find "Hari_Masuk = '" & textKriteria.Text & "'"
        Case Is = 1
            .Find "Tanggal_Masuk = '" & textKriteria.Text & "'"
        Case Is = 2
            .Find "Bulan_Masuk = '" & textKriteria.Text & "'"
        Case Is = 3
            .Find "Tahun_Masuk = '" & textKriteria.Text & "'"
        Case Is = 4
            .Find "Pengirim = '" & textKriteria.Text & "'"
        Case Is = 5
            .Find "Jenis = '" & textKriteria.Text & "'"
        Case Is = 6
            .Find "Kode = '" & textKriteria.Text & "'"
        Case Is = 7
            .Find "Judul = '" & textKriteria.Text & "'"
        Case Is = 8
            .Find "Kategori = '" & textKriteria.Text & "'"
        Case Is = 9
            .Find "Pengarang = '" & textKriteria.Text & "'"
        Case Is = 10
            .Find "Penerbit = '" & textKriteria.Text & "'"
        Case Is = 11
            .Find "Tahun_Terbit = '" & textKriteria.Text & "'"
        Case Is = 12
            .Find "Cetakan_Ke = '" & textKriteria.Text & "'"
        Case Is = 13
            .Find "Jumlah = '" & textKriteria.Text & "'"
        Case Is = 14
            .Find "Stok = '" & textKriteria.Text & "'"
        Case Is = 15
            .Find "Keterangan = '" & textKriteria.Text & "'"
        End Select
        If .EOF Then
            MsgBox ("Data tidak ditemukan"), vbExclamation + vbOKOnly, ""
        Else
            Set FormManageBukuMasuk.DataGrid1.DataSource = FormManageBukuMasuk.AdodcUtama.Recordset
            cmBatal.Caption = "&Tutup"
            If FormPengaturan.cekTutupFormCari.Value = Checked Then Me.Hide
        End If
    End With
End If
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub



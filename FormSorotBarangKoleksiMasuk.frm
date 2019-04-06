VERSION 5.00
Begin VB.Form FormSorotBarangKoleksiMasuk 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sorot Data - Barang Koleksi Masuk"
   ClientHeight    =   1785
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
   Icon            =   "FormSorotBarangKoleksiMasuk.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   6240
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
      Width           =   1335
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
      Width           =   6015
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
         TabIndex        =   3
         Top             =   720
         Width           =   1335
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
         TabIndex        =   2
         Top             =   720
         Width           =   2295
      End
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
         TabIndex        =   1
         Top             =   240
         Width           =   3735
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
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FormSorotBarangKoleksiMasuk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    With cmbSorotDataBerdasarkan
        .Clear
        .AddItem FormManageBarangKoleksiMasuk.AdodcUtama.Recordset.Fields(0).Name, 0
        .AddItem FormManageBarangKoleksiMasuk.AdodcUtama.Recordset.Fields(1).Name, 1
        .AddItem FormManageBarangKoleksiMasuk.AdodcUtama.Recordset.Fields(2).Name, 2
        .AddItem FormManageBarangKoleksiMasuk.AdodcUtama.Recordset.Fields(3).Name, 3
        .AddItem FormManageBarangKoleksiMasuk.AdodcUtama.Recordset.Fields(4).Name, 4
        .AddItem FormManageBarangKoleksiMasuk.AdodcUtama.Recordset.Fields(5).Name, 5
        .AddItem FormManageBarangKoleksiMasuk.AdodcUtama.Recordset.Fields(6).Name, 6
        .AddItem FormManageBarangKoleksiMasuk.AdodcUtama.Recordset.Fields(7).Name, 7
        .AddItem FormManageBarangKoleksiMasuk.AdodcUtama.Recordset.Fields(8).Name, 8
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
    With FormManageBarangKoleksiMasuk
        .AdodcUtama.Refresh
            Select Case cmbSorotDataBerdasarkan.ListIndex
            Case Is = 0
                .AdodcUtama.RecordSource = "Select * from tbBarangKoleksiMasuk order by Hari asc;"
            Case Is = 1
                .AdodcUtama.RecordSource = "Select * from tbBarangKoleksiMasuk order by Tanggal asc;"
            Case Is = 2
                .AdodcUtama.RecordSource = "Select * from tbBarangKoleksiMasuk order by Bulan asc;"
            Case Is = 3
                .AdodcUtama.RecordSource = "Select * from tbBarangKoleksiMasuk order by Tahun asc;"
            Case Is = 4
                .AdodcUtama.RecordSource = "Select * from tbBarangKoleksiMasuk order by Kode_Barang asc;"
            Case Is = 5
                .AdodcUtama.RecordSource = "Select * from tbBarangKoleksiMasuk order by Nama_Barang asc;"
            Case Is = 6
                .AdodcUtama.RecordSource = "Select * from tbBarangKoleksiMasuk order by Status_Barang asc;"
            Case Is = 7
                .AdodcUtama.RecordSource = "Select * from tbBarangKoleksiMasuk order by Penerima_(Nama_Admin) asc;"
            Case Is = 8
                .AdodcUtama.RecordSource = "Select * from tbBarangKoleksiMasuk order by Keterangan asc;"
            End Select
    End With
ElseIf cmbMode.ListIndex = 1 Then
    With FormManageBarangKoleksiMasuk
        .AdodcUtama.Refresh
            Select Case cmbSorotDataBerdasarkan.ListIndex
            Case Is = 0
                .AdodcUtama.RecordSource = "Select * from tbBarangKoleksiMasuk order by Hari desc;"
            Case Is = 1
                .AdodcUtama.RecordSource = "Select * from tbBarangKoleksiMasuk order by Tanggal desc;"
            Case Is = 2
                .AdodcUtama.RecordSource = "Select * from tbBarangKoleksiMasuk order by Bulan desc;"
            Case Is = 3
                .AdodcUtama.RecordSource = "Select * from tbBarangKoleksiMasuk order by Tahun desc;"
            Case Is = 4
                .AdodcUtama.RecordSource = "Select * from tbBarangKoleksiMasuk order by Kode_Barang desc;"
            Case Is = 5
                .AdodcUtama.RecordSource = "Select * from tbBarangKoleksiMasuk order by Nama_Barang desc;"
            Case Is = 6
                .AdodcUtama.RecordSource = "Select * from tbBarangKoleksiMasuk order by Status_Barang desc;"
            Case Is = 7
                .AdodcUtama.RecordSource = "Select * from tbBarangKoleksiMasuk order by Penerima_(Nama_Admin) desc;"
            Case Is = 8
                .AdodcUtama.RecordSource = "Select * from tbBarangKoleksiMasuk order by Keterangan desc;"
        End Select
    End With
End If
    FormManageBarangKoleksiMasuk.AdodcUtama.Refresh
    cmBatal.Caption = "&Tutup"
    If FormPengaturan.cekTutupFormSorot.Value = Checked Then Me.Hide
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub



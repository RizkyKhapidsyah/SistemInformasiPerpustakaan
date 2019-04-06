VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormManageBarangKoleksiMasuk 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Data - Barang Koleksi Masuk"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12885
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormManageBarangKoleksiMasuk.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   12885
   Begin VB.CommandButton cmNavigasiDataAkhir 
      BackColor       =   &H00E0E0E0&
      Caption         =   ">>"
      Height          =   495
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton cmNavigasiDataSelanjutnya 
      BackColor       =   &H00E0E0E0&
      Caption         =   ">"
      Height          =   495
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton cmNavigasiDataSebelumnya 
      BackColor       =   &H00E0E0E0&
      Caption         =   "<"
      Height          =   495
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton cmNavigasiDataAwal 
      BackColor       =   &H00E0E0E0&
      Caption         =   "<<"
      Height          =   495
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4560
      Width           =   495
   End
   Begin MSAdodcLib.Adodc AdodcUtama 
      Height          =   330
      Left            =   10320
      Top             =   360
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
   Begin VB.CommandButton cmTambah 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Tambah"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmEdit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Edit"
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmCari 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Cari"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmSorot 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Sorot"
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmFilter 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Filter"
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmHapus 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmTutup 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Tutup"
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmRefresh 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Refresh"
      Height          =   495
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   7646
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   20
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FormManageBarangKoleksiMasuk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    NyambunggUtama
    With AdodcUtama
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * from tbBarangKoleksiMasuk order by NamaBarang asc;"
        Set DataGrid1.DataSource = AdodcUtama
        .Refresh
    End With
    DataGrid1.AllowUpdate = False
    With DataGrid1
        .Columns(0).Width = 434.8347
        .Columns(1).Width = 615.1182
        .Columns(2).Width = 464.8819
        .Columns(3).Width = 540.2835
        .Columns(4).Width = 1170.142
        .Columns(5).Width = 2355.024
        .Columns(6).Width = 1739.906
        .Columns(7).Width = 2355.024
        .Columns(8).Width = 2550.047
    End With
End Sub

Private Sub cmCari_Click()
If AdodcUtama.Recordset.RecordCount = 0 Then
    MsgBox "Tidak ada data yang akan dicari!", vbExclamation + vbOKOnly, ""
Else
    With FormCariBarangKoleksi
        .Show
        .SetFocus
    End With
End If
End Sub

Private Sub cmEdit_Click()
If AdodcUtama.Recordset.RecordCount = 0 Then
    MsgBox "Tidak ada data yang akan diedit!", vbExclamation + vbOKOnly, ""
Else
    With FormBarangKoleksiMasuk
        .Caption = "Edit Data - Barang Koleksi Masuk"
        .AturKontrol
        
        .cmbHari.Text = AdodcUtama.Recordset.Fields(0).Value
        .cmbTanggal.Text = AdodcUtama.Recordset.Fields(1).Value
        .cmbBulan.Text = AdodcUtama.Recordset.Fields(2).Value
        .cmbTahun.Text = AdodcUtama.Recordset.Fields(3).Value
        .textKodeBarang.Text = AdodcUtama.Recordset.Fields(4).Value
        .textNamaBarang.Text = AdodcUtama.Recordset.Fields(5).Value
        .cmbStatusBarang.Text = AdodcUtama.Recordset.Fields(6).Value
        .textPenerima.Text = AdodcUtama.Recordset.Fields(7).Value
        .textKeterangan.Text = AdodcUtama.Recordset.Fields(8).Value
        
        .cmSimpan.Caption = "&Perbarui"
        .cmTutup.Caption = "&Batal"
        .Show
        .SetFocus
    End With
End If
End Sub

Private Sub cmFilter_Click()
If AdodcUtama.Recordset.RecordCount = 0 Then
    MsgBox "Tidak ada data yang akan disorot!", vbExclamation + vbOKOnly, ""
Else
    With FormFilterBarangKoleksiMasuk
        .Show
        .SetFocus
    End With
End If
End Sub

Private Sub cmHapus_Click()
If AdodcUtama.Recordset.RecordCount = 0 Then
    MsgBox "Tidak ada data yang dapat dihapus", vbExclamation + vbOKOnly, ""
Else
    X = MsgBox("Apakah Anda yakin ingin menghapus data ini?", vbQuestion + vbYesNo, "Hapus data?")
    If X = vbYes Then
        With AdodcUtama
            .Recordset.Delete
            .Refresh
        End With
    End If
End If
End Sub

Private Sub cmNavigasiDataAkhir_Click()
    AdodcUtama.Recordset.MoveLast
End Sub

Private Sub cmNavigasiDataAwal_Click()
    AdodcUtama.Recordset.MoveFirst
End Sub

Private Sub cmNavigasiDataSebelumnya_Click()
        AdodcUtama.Recordset.MovePrevious
    If AdodcUtama.Recordset.BOF = True Then AdodcUtama.Recordset.MoveLast
End Sub

Private Sub cmNavigasiDataSelanjutnya_Click()
        AdodcUtama.Recordset.MoveNext
    If AdodcUtama.Recordset.EOF = True Then AdodcUtama.Recordset.MoveFirst
End Sub

Private Sub cmRefresh_Click()
    AturKontrol
    With Me
        .cmEdit.Enabled = True
        .cmCari.Enabled = True
        .cmSorot.Enabled = True
        .cmFilter.Enabled = True
        .cmHapus.Enabled = True
        .cmTambah.Enabled = True
    End With
End Sub

Private Sub cmSorot_Click()
If AdodcUtama.Recordset.RecordCount = 0 Then
    MsgBox "Tidak ada data yang akan disorot!", vbExclamation + vbOKOnly, ""
Else
    With FormSorotBarangKoleksiMasuk
        .Show
        .SetFocus
    End With
End If
End Sub

Private Sub cmTambah_Click()
With FormBarangKoleksiMasuk
    .Caption = "Tambah Data - Barang Koleksi Masuk"
    .AturKontrol
    .Show
    .SetFocus
End With
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub


Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    AturKontrol
End Sub

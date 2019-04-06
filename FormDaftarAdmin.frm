VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormDaftarAdmin 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daftar Admin"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3855
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormDaftarAdmin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   3855
   Begin VB.TextBox textNamaPengguna 
      Height          =   390
      Left            =   120
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   6015
      Left            =   1680
      TabIndex        =   9
      Top             =   0
      Width           =   2295
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5295
         Left            =   0
         TabIndex        =   10
         Top             =   120
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   9340
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
   Begin VB.CommandButton cmAkhir 
      BackColor       =   &H00E0E0E0&
      Caption         =   ">>"
      Height          =   495
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4920
      Width           =   360
   End
   Begin VB.CommandButton CmSelanjutnya 
      BackColor       =   &H00E0E0E0&
      Caption         =   ">"
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4920
      Width           =   360
   End
   Begin VB.CommandButton cmSebelumnya 
      BackColor       =   &H00E0E0E0&
      Caption         =   "<"
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Width           =   360
   End
   Begin VB.CommandButton cmAwal 
      BackColor       =   &H00E0E0E0&
      Caption         =   "<<"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
      Width           =   360
   End
   Begin VB.CommandButton cmTutup 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Tutup"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   1440
   End
   Begin VB.CommandButton cmRefresh 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Refresh"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   1440
   End
   Begin VB.CommandButton cmHapus 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1440
   End
   Begin VB.CommandButton cmEdit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Edit"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   1440
   End
   Begin VB.CommandButton cmTambah 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Tambah"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1440
   End
   Begin MSAdodcLib.Adodc AdodcUtamaPWTampilkan 
      Height          =   450
      Left            =   8760
      Top             =   5760
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   794
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
Attribute VB_Name = "FormDaftarAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
NyambunggUtama
With AdodcUtamaPWTampilkan
    .ConnectionString = CN.ConnectionString
    .RecordSource = "Select * From tbLogin;"
    Set DataGrid1.DataSource = AdodcUtamaPWTampilkan
    .Refresh
End With
DataGrid1.AllowUpdate = False
textNamaPengguna.Text = AdodcUtamaPWTampilkan.Recordset.Fields(0).Value
End Sub

Private Sub cmAkhir_Click()
    AdodcUtamaPWTampilkan.Recordset.MoveLast
    textNamaPengguna.Text = AdodcUtamaPWTampilkan.Recordset.Fields(0).Value
End Sub

Private Sub cmAwal_Click()
    AdodcUtamaPWTampilkan.Recordset.MoveFirst
    textNamaPengguna.Text = AdodcUtamaPWTampilkan.Recordset.Fields(0).Value
End Sub

Private Sub cmEdit_Click()
With FormAutorisasi
    .Caption = "Autentikasi - (Admin - @" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & ")"
    .LabelJenisAksi.Caption = cmEdit.Caption
    .textAutentikasiSembunyi.Text = AdodcUtamaPWTampilkan.Recordset.Fields(0).Value
    .Caption = "Edit Data - (Admin@" & AdodcUtamaPWTampilkan.Recordset.Fields(0).Value & ")"
    .Show
    .SetFocus
End With
End Sub

Private Sub cmHapus_Click()
If textNamaPengguna.Text = FORM_UTAMA.StatusBawah.Panels.Item(2).Text Then
    MsgBox "Tidak dapat menghapus diri Anda dari sistem!", vbCritical + vbOKOnly, "Main System - Akses Gagal"
Else
    With FormAutorisasi
        .LabelJenisAksi.Caption = cmHapus.Caption
        .textAutentikasiSembunyi.Text = AdodcUtamaPWTampilkan.Recordset.Fields(0).Value
        .Caption = "Hapus Data - (Admin@" & AdodcUtamaPWTampilkan.Recordset.Fields(0).Value & ")"
        .Label1.Caption = "Masukkan Password"
        .Show
        .SetFocus
    End With
End If
End Sub

Private Sub cmRefresh_Click()
    AturKontrol
End Sub

Private Sub cmSebelumnya_Click()
    AdodcUtamaPWTampilkan.Recordset.MovePrevious
        If AdodcUtamaPWTampilkan.Recordset.BOF = True Then AdodcUtamaPWTampilkan.Recordset.MoveLast
    textNamaPengguna.Text = AdodcUtamaPWTampilkan.Recordset.Fields(0).Value
End Sub

Private Sub CmSelanjutnya_Click()
    AdodcUtamaPWTampilkan.Recordset.MoveNext
        If AdodcUtamaPWTampilkan.Recordset.EOF = True Then AdodcUtamaPWTampilkan.Recordset.MoveFirst
    textNamaPengguna.Text = AdodcUtamaPWTampilkan.Recordset.Fields(0).Value
End Sub

Private Sub cmTambah_Click()
With FormAutorisasi
    .Caption = "Autentikasi - (Admin - @" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & ")"
    .LabelJenisAksi.Caption = cmTambah.Caption
    .textAutentikasiSembunyi.Text = FORM_UTAMA.StatusBawah.Panels.Item(2).Text
    .Caption = "Autorisasi - (Admin@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & ")"
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

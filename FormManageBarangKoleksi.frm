VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormManageBarangKoleksi 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage - Barang Koleksi"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9735
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormManageBarangKoleksi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   9735
   Begin MSAdodcLib.Adodc AdodcUtama 
      Height          =   375
      Left            =   360
      Top             =   5880
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
   Begin VB.CommandButton cmRefresh 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Refresh"
      Height          =   495
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmTutup 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Tutup"
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmHapus 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmFilter 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Filter"
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmSorot 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Sorot"
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmCari 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Cari"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmEdit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Edit"
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmBaru 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Baru"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   7435
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
Attribute VB_Name = "FormManageBarangKoleksi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    NyambunggUtama
    With AdodcUtama
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * from tbBarangKoleksi order by Kode_Barang asc;"
        Set DataGrid1.DataSource = AdodcUtama
        .Refresh
    End With
End Sub

Private Sub cmBaru_Click()
    With formBarangKoleksi
        .AturKontrol
        .Show
        .SetFocus
    End With
End Sub

Private Sub cmCari_Click()
    If AdodcUtama.Recordset.RecordCount = 0 Then
        MsgBox "Tidak ada data barang yang akan dicari!", vbExclamation + vbOKOnly, ""
    Else
        With FormCariDataKoleksiBarang
            .Show
            .SetFocus
        End With
    End If
End Sub

Private Sub cmEdit_Click()
If AdodcUtama.Recordset.RecordCount = 0 Then
    MsgBox "Tidak ada data yang akan diedit", vbExclamation + vbOKOnly, ""
Else
    With formBarangKoleksi
        .textKodeBarang.Text = AdodcUtama.Recordset.Fields(0).Value
        .textNamaBarang.Text = AdodcUtama.Recordset.Fields(1).Value
        .cmbKepemilikan.Text = AdodcUtama.Recordset.Fields(2).Value
        .TextKeterangan.Text = AdodcUtama.Recordset.Fields(3).Value
        .Caption = "Edit Data - Barang Koleksi"
        .cmSimpan.Caption = "&Perbarui"
        .Show
        .SetFocus
    End With
End If
End Sub

Private Sub cmFilter_Click()
    If AdodcUtama.Recordset.RecordCount = 0 Then
        MsgBox "Tidak ada data barang yang akan disorot!", vbExclamation + vbOKOnly, ""
    Else
        With FormFilterBarangKoleksi
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

Private Sub cmRefresh_Click()
    Form_Load
    AturKontrol
    cmEdit.Enabled = True
    cmCari.Enabled = True
    cmSorot.Enabled = True
    cmFilter.Enabled = True
    cmHapus.Enabled = True
    cmBaru.Enabled = True
End Sub

Private Sub cmSorot_Click()
    If AdodcUtama.Recordset.RecordCount = 0 Then
        MsgBox "Tidak ada data barang yang akan disorot!", vbExclamation + vbOKOnly, ""
    Else
        With FormSorotBarangKoleksi
            .Show
            .SetFocus
        End With
    End If
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub

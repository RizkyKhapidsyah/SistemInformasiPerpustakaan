VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.ocx"
Begin VB.Form FormAnggota 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anggota"
   ClientHeight    =   4095
   ClientLeft      =   3060
   ClientTop       =   5025
   ClientWidth     =   12390
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormAnggota.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   12390
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmRefresh 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Refresh"
      Height          =   495
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmTutup 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Tutup"
      Height          =   495
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmHapus 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmFilter 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Filter"
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmSorot 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Sorot"
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmCari 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Cari"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmEdit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Edit"
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmTambah 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Tambah"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   4560
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   5741
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
Attribute VB_Name = "FormAnggota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    NyambunggUtama
    With Adodc1
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * from tbAnggota"
        Set DataGrid1.DataSource = Adodc1
        .Refresh
    End With
    With FormPengaturan
        If .cekKunciTabel.Value = Checked Then
            DataGrid1.AllowUpdate = False
        ElseIf .cekKunciTabel.Value = Unchecked Then
            DataGrid1.AllowUpdate = True
        End If
    End With
    If FORM_UTAMA.StatusBawah.Panels.Item(2).Text = "Anonymous" Then
        With Me
            .cmTambah.Enabled = False
            .cmEdit.Enabled = False
            .cmHapus.Enabled = False
        End With
    End If
    With DataGrid1
        .Columns(0).Width = 1065.26
        .Columns(1).Width = 1590.236
        .Columns(2).Width = 1214.929
        .Columns(3).Width = 1184.882
        .Columns(4).Width = 975.1182
        .Columns(5).Width = 854.9292
    End With
End Sub

Private Sub cmCari_Click()
If Adodc1.Recordset.RecordCount = 0 Then
    MsgBox "Tidak ada data yang dapat dicari", vbExclamation + vbOKOnly, ""
Else
    With FormCari
        .Show
        .SetFocus
    End With
End If
End Sub

Private Sub cmEdit_Click()
If Adodc1.Recordset.RecordCount = 0 Then
    MsgBox "Tidak ada data yang dapat diedit", vbExclamation + vbOKOnly, ""
Else
    With FormTambahAnggota
        .textNomorPokok.Text = Adodc1.Recordset.Fields(0).Value
        .textNama.Text = Adodc1.Recordset.Fields(1).Value
        .cmbJenisKelamin.Text = Adodc1.Recordset.Fields(2).Value
        .TextNomorHP.Text = Adodc1.Recordset.Fields(3).Value
        .textAlamat.Text = Adodc1.Recordset.Fields(4).Value
        .cmbStatus.Text = Adodc1.Recordset.Fields(5).Value
        .cmSImpan.Caption = "&Update"
        .Caption = "Edit Anggota"
        .Show
        .SetFocus
    End With
End If
End Sub

Private Sub cmFilter_Click()
If Adodc1.Recordset.RecordCount = 0 Then
    MsgBox "Tidak ada data yang dapat difilter", vbExclamation + vbOKOnly, ""
Else
    With FormFilterAnggota
        .Show
        .SetFocus
    End With
End If
End Sub

Private Sub cmHapus_Click()
If Adodc1.Recordset.RecordCount = 0 Then
    MsgBox "Tidak ada data yang dapat dihapus", vbExclamation + vbOKOnly, ""
Else
    X = MsgBox("Apakah Anda yakin ingin menghapus data ini?" & vbCrLf & vbCrLf & _
            "------------------------------------------" & vbCrLf & _
            "Nomor Pokok : " & Adodc1.Recordset.Fields(0).Value & vbCrLf & _
            "Nama : " & Adodc1.Recordset.Fields(1).Value & vbCrLf & _
            "Jenis Kelamin : " & Adodc1.Recordset.Fields(2).Value & vbCrLf & _
            "Nomor HP : " & Adodc1.Recordset.Fields(3).Value & vbCrLf & _
            "Alamat : " & Adodc1.Recordset.Fields(4).Value & vbCrLf & _
            "Status : " & Adodc1.Recordset.Fields(5).Value & vbCrLf & _
            "------------------------------------------" & vbCrLf & vbCrLf & _
            "Yakin ingin menghapus data ini?", vbQuestion + vbYesNo, "Hapus data?")
    If X = vbYes Then
        With Adodc1
            .Recordset.Delete
            .Refresh
        End With
    End If
End If
End Sub

Public Sub cmRefresh_Click()
    AturKontrol
    If FORM_UTAMA.StatusBawah.Panels.Item(2).Text = "Anonymous" Then
        cmTambah.Enabled = False
        cmEdit.Enabled = False
        cmHapus.Enabled = False
        cmCari.Enabled = True
        cmSorot.Enabled = True
        cmFilter.Enabled = True
    Else
        cmEdit.Enabled = True
        cmCari.Enabled = True
        cmSorot.Enabled = True
        cmFilter.Enabled = True
        cmHapus.Enabled = True
    End If
End Sub

Private Sub cmSorot_Click()
If Adodc1.Recordset.RecordCount = 0 Then
    MsgBox "Tidak ada data yang dapat disorot", vbExclamation + vbOKOnly, ""
Else
    With FormSorotAnggota
        .Show
        .SetFocus
    End With
End If
End Sub

Private Sub cmTambah_Click()
    cmRefresh_Click
    With FormTambahAnggota
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

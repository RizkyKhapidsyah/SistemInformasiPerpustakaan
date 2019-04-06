VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormBarangKoleksiMasuk 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Barang Koleksi Masuk"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6840
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormBarangKoleksiMasuk.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6840
   Begin MSAdodcLib.Adodc AdodcUtama 
      Height          =   330
      Left            =   4440
      Top             =   4080
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
   Begin VB.CommandButton cmTutup 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Tutup"
      Height          =   495
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmReset 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Reset"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmManage 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Manage"
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmSimpan 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.CommandButton cmSetTanggal 
         BackColor       =   &H00E0E0E0&
         Caption         =   "<"
         Height          =   375
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Set Tanggal Saat Ini"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmAcakKode 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Acak Kode"
         Height          =   375
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox textKeterangan 
         Height          =   1095
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Text            =   "FormBarangKoleksiMasuk.frx":000C
         Top             =   2640
         Width           =   4575
      End
      Begin VB.TextBox textPenerima 
         Height          =   390
         Left            =   1920
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   2160
         Width           =   2415
      End
      Begin VB.ComboBox cmbStatusBarang 
         Height          =   390
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox textNamaBarang 
         Height          =   390
         Left            =   1920
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1200
         Width           =   3855
      End
      Begin VB.TextBox textKodeBarang 
         Height          =   390
         Left            =   1920
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   720
         Width           =   2415
      End
      Begin VB.ComboBox cmbTahun 
         Height          =   390
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmbBulan 
         Height          =   390
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cmbTanggal 
         Height          =   390
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cmbHari 
         Height          =   390
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   21
         Top             =   2640
         Width           =   45
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         Height          =   270
         Left            =   120
         TabIndex        =   20
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   18
         Top             =   2160
         Width           =   45
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Penerima (Admin)"
         Height          =   270
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   1170
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   15
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Barang"
         Height          =   270
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   915
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   12
         Top             =   1200
         Width           =   45
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Barang"
         Height          =   270
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   870
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   9
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Barang"
         Height          =   270
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   270
         Left            =   3360
         TabIndex        =   4
         Top             =   240
         Width           =   60
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Masuk"
         Height          =   270
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   930
      End
   End
End
Attribute VB_Name = "FormBarangKoleksiMasuk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    With AdodcUtama
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * from tbBarangKoleksiMasuk order by NamaBarang asc;"
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
    With cmbStatusBarang
        .Clear
        .AddItem "Internal BPAD Provsu", 0
        .AddItem "Eksternal BPAD Provsu", 1
        .AddItem "Sewa", 2
        .AddItem "Beli", 3
        .ListIndex = 0
    End With
    KosongkanInput
End Sub
Sub KosongkanInput()
    For Each Objek In Me
        If TypeName(Objek) = "TextBox" Then
            With Objek
                .Text = ""
                .MaxLength = 254
            End With
        End If
    Next
    With textKodeBarang
        .Text = "IV" & Second(Time) & "BPAD" & Minute(Time) & "T" & Hour(Time)
        .Locked = True
    End With
End Sub

Private Sub cmbBulan_Click()
    cmbTanggal.Clear
    Select Case cmbBulan.ListIndex
        Case Is = 0, 2, 4, 6, 7, 9, 11
            For X = 1 To 31
                With cmbTanggal
                    .AddItem X
                End With
            Next
        Case Is = 1
            If Val(cmbTahun.Text) Mod 4 Then
                For X = 1 To 29
                    With cmbTanggal
                        .AddItem X
                    End With
                Next
            Else
                For X = 1 To 28
                    With cmbTanggal
                        .AddItem X
                    End With
                Next
            End If
        Case Is = 3, 5, 8, 10
            For X = 1 To 30
                With cmbTanggal
                    .AddItem X
                End With
            Next
    End Select
    cmbTanggal.Text = "1"
End Sub

Private Sub cmManage_Click()
With FormManageBarangKoleksiMasuk
    .Show
    .SetFocus
End With
End Sub

Private Sub cmReset_Click()
    KosongkanInput
    textNamaBarang.SetFocus
End Sub

Private Sub cmSetTanggal_Click()
    cmbTahun.Text = Year(Date)
    cmbBulan.Text = Month(Date)
    cmbTanggal.Text = Day(Date)
End Sub

Private Sub cmSimpan_Click()
On Error GoTo HancurkanError
    If textNamaBarang.Text = "" Then
        MsgBox "Silahkan isi nama barang yang akan diinput", vbExclamation + vbOKOnly, ""
        textNamaBarang.SetFocus
    ElseIf textPenerima.Text = "" Then
        MsgBox "Silahkan isi nama Admin yang menerima barang (Nama Anda)", vbExclamation + vbOKOnly, ""
        textPenerima.SetFocus
    ElseIf textKeterangan.Text = "" Then
        MsgBox "Silahkan isi keterangan yang diperlukan", vbExclamation + vbOKOnly, ""
        textKeterangan.SetFocus
    Else
        Select Case cmSimpan.Caption
        Case "&Simpan"
            X = MsgBox("Anda yakin ingin menyimpan data ini?", vbQuestion + vbYesNo, "Konfirmasi")
                If X = vbYes Then
                    With AdodcUtama
                        .Recordset.AddNew
                        .Recordset.Fields(0).Value = cmbHari.Text
                        .Recordset.Fields(1).Value = cmbTanggal.Text
                        .Recordset.Fields(2).Value = cmbBulan.Text
                        .Recordset.Fields(3).Value = cmbTahun.Text
                        .Recordset.Fields(4).Value = textKodeBarang.Text
                        .Recordset.Fields(5).Value = textNamaBarang.Text
                        .Recordset.Fields(6).Value = cmbStatusBarang.Text
                        .Recordset.Fields(7).Value = textPenerima.Text
                        .Recordset.Fields(8).Value = textKeterangan.Text
                        .Recordset.Update
                        .Refresh
                    End With
                    KosongkanInput
                    FormManageBarangKoleksiMasuk.AturKontrol
                End If
        Case "&Perbarui"
            X = MsgBox("Anda yakin ingin memperbarui data ini?", vbQuestion + vbYesNo, "Konfirmasi")
                If X = vbYes Then
                    With FormManageBarangKoleksiMasuk.AdodcUtama
                        .Recordset.Delete
                        .Recordset.AddNew
                        .Recordset.Fields(0).Value = cmbHari.Text
                        .Recordset.Fields(1).Value = cmbTanggal.Text
                        .Recordset.Fields(2).Value = cmbBulan.Text
                        .Recordset.Fields(3).Value = cmbTahun.Text
                        .Recordset.Fields(4).Value = textKodeBarang.Text
                        .Recordset.Fields(5).Value = textNamaBarang.Text
                        .Recordset.Fields(6).Value = cmbStatusBarang.Text
                        .Recordset.Fields(7).Value = textPenerima.Text
                        .Recordset.Fields(8).Value = textKeterangan.Text
                        .Recordset.Update
                        .Refresh
                    End With
                    KosongkanInput
                    FormManageBarangKoleksiMasuk.AturKontrol
                    cmSimpan.Caption = "&Simpan"
                    Me.Caption = "Tambah Data - Barang Koleksi Masuk"
                End If
        End Select
        textNamaBarang.SetFocus
    End If
Exit Sub
HancurkanError:
    PusatError
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub

Private Sub cmAcakKode_Click()
    With textKodeBarang
        .Text = "IV" & Second(Time) & "BPAD" & Minute(Time) & "T" & Hour(Time)
        .Locked = True
    End With
    textNamaBarang.SetFocus
End Sub

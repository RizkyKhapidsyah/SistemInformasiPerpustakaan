VERSION 5.00
Begin VB.Form FormTambahAnggota 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tambah Anggota"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormTambahAnggota.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmManage 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Manage"
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   6615
      Begin VB.TextBox textNomorPokok 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   2520
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox textNama 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   2520
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   720
         Width           =   3975
      End
      Begin VB.TextBox TextNomorHP 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   2520
         TabIndex        =   6
         Top             =   1680
         Width           =   2775
      End
      Begin VB.ComboBox cmbJenisKelamin 
         Height          =   390
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox textAlamat 
         Appearance      =   0  'Flat
         Height          =   855
         Left            =   2520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Text            =   "FormTambahAnggota.frx":000C
         Top             =   2160
         Width           =   3975
      End
      Begin VB.ComboBox cmbStatus 
         Height          =   390
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   3120
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIA"
         Height          =   270
         Left            =   2085
         TabIndex        =   20
         Top             =   240
         Width           =   210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2400
         TabIndex        =   19
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   270
         Left            =   1935
         TabIndex        =   18
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2400
         TabIndex        =   17
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Kelamin"
         Height          =   270
         Left            =   1410
         TabIndex        =   16
         Top             =   1200
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2400
         TabIndex        =   15
         Top             =   1200
         Width           =   45
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor HP"
         Height          =   270
         Left            =   1635
         TabIndex        =   14
         Top             =   1680
         Width           =   660
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2400
         TabIndex        =   13
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         Height          =   270
         Left            =   1860
         TabIndex        =   12
         Top             =   2160
         Width           =   435
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2400
         TabIndex        =   11
         Top             =   2160
         Width           =   45
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   270
         Left            =   1920
         TabIndex        =   10
         Top             =   3120
         Width           =   405
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2430
         TabIndex        =   9
         Top             =   3120
         Width           =   45
      End
   End
   Begin VB.CommandButton cmTutup 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Batal"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmSImpan 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      Width           =   1095
   End
End
Attribute VB_Name = "FormTambahAnggota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
For Each Objek In Me
    If TypeName(Objek) = "TextBox" Then
        With Objek
            .Text = ""
            .MaxLength = 254
        End With
    End If
Next
'LAINNYA
With cmbJenisKelamin
    .Clear
    .AddItem "Laki-Laki", 0
    .AddItem "Perempuan", 1
    .ListIndex = 0
End With
With cmbStatus
    .Clear
    .AddItem "Admin", 0
    .AddItem "Pegawai/Staff", 1
    .AddItem "Umum", 2
    .ListIndex = 0
End With
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
End Sub

Private Sub cmManage_Click()
With FormAnggota
    .Show
    .SetFocus
End With
End Sub

Private Sub cmSimpan_Click()
    If textNomorPokok.Text = "" Then
        MsgBox "Silahkan isi Nomor Pokok Anda!", vbExclamation + vbOKOnly, ""
        textNomorPokok.SetFocus
    ElseIf textNama.Text = "" Then
        MsgBox "Silahkan isi Nama Anda!", vbExclamation + vbOKOnly, ""
        textNama.SetFocus
    ElseIf TextNomorHP.Text = "" Then
        MsgBox "Silahkan isi Nomor HP Anda!", vbExclamation + vbOKOnly, ""
        TextNomorHP.SetFocus
    ElseIf textAlamat.Text = "" Then
        MsgBox "Silahkan isi Alamat tempat tinggal Anda!", vbExclamation + vbOKOnly, ""
        textAlamat.SetFocus
    Else
        X = MsgBox("Anda yakin ingin menyimpan data ini?", vbQuestion + vbYesNo, "Konfirmasi")
        If X = vbYes Then
            If cmSImpan.Caption = "&Simpan" Then
                With FormAnggota.Adodc1
                    .Recordset.AddNew
                    .Recordset.Fields(0).Value = textNomorPokok.Text
                    .Recordset.Fields(1).Value = textNama.Text
                    .Recordset.Fields(2).Value = cmbJenisKelamin.Text
                    .Recordset.Fields(3).Value = TextNomorHP.Text
                    .Recordset.Fields(4).Value = textAlamat.Text
                    .Recordset.Fields(5).Value = cmbStatus.Text
                    .Recordset.Update
                    .Refresh
                End With
                KosongkanInput
                MsgBox "Data Berhasil disimpan!", vbInformation + vbOKOnly, "Berhasil"
                cmTutup.Caption = "&Tutup"
            ElseIf cmSImpan.Caption = "&Update" Then
                With FormAnggota.Adodc1
                    .Recordset.Delete
                    .Recordset.AddNew
                    .Recordset.Fields(0).Value = textNomorPokok.Text
                    .Recordset.Fields(1).Value = textNama.Text
                    .Recordset.Fields(2).Value = cmbJenisKelamin.Text
                    .Recordset.Fields(3).Value = TextNomorHP.Text
                    .Recordset.Fields(4).Value = textAlamat.Text
                    .Recordset.Fields(5).Value = cmbStatus.Text
                    .Recordset.Update
                    .Refresh
                End With
                KosongkanInput
                MsgBox "Data Berhasil diedit dan disimpan!", vbInformation + vbOKOnly, "Berhasil"
                cmTutup.Caption = "&Tutup"
                cmSImpan.Caption = "&Simpan"
                Me.Caption = "Tambah Anggota"
            End If
        End If
    End If
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub


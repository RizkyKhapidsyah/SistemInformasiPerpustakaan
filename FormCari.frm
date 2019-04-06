VERSION 5.00
Begin VB.Form FormCari 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cari Data "
   ClientHeight    =   2280
   ClientLeft      =   7830
   ClientTop       =   6375
   ClientWidth     =   3240
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormCari.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmBatal 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Batal"
      Height          =   375
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmCari 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Cari"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox textKriteria 
      Height          =   390
      Left            =   120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1320
      Width           =   3015
   End
   Begin VB.ComboBox cmbCariDataBerdasarkan 
      Height          =   390
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dengan Kriteria :"
      Height          =   270
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cari Data Berdasarkan :"
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "FormCari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
With Me
    .Caption = "Cari Data - Anggota"
    .cmbCariDataBerdasarkan.Clear
    .cmbCariDataBerdasarkan.AddItem FormAnggota.Adodc1.Recordset.Fields(0).Name, 0
    .cmbCariDataBerdasarkan.AddItem FormAnggota.Adodc1.Recordset.Fields(1).Name, 1
    .cmbCariDataBerdasarkan.AddItem FormAnggota.Adodc1.Recordset.Fields(2).Name, 2
    .cmbCariDataBerdasarkan.AddItem FormAnggota.Adodc1.Recordset.Fields(3).Name, 3
    .cmbCariDataBerdasarkan.AddItem FormAnggota.Adodc1.Recordset.Fields(4).Name, 4
    .cmbCariDataBerdasarkan.AddItem FormAnggota.Adodc1.Recordset.Fields(5).Name, 5
    .cmbCariDataBerdasarkan.ListIndex = 0
    .textKriteria.Text = ""
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
    FormAnggota.Adodc1.Refresh
    With FormAnggota.Adodc1.Recordset
        Select Case cmbCariDataBerdasarkan.ListIndex
        Case Is = 0
            .Find "NIA = '" & textKriteria.Text & "'"
        Case Is = 1
            .Find "Nama = '" & textKriteria.Text & "'"
        Case Is = 2
            .Find "JenisKelamin = '" & textKriteria.Text & "'"
        Case Is = 3
            .Find "Nomor_HP = '" & textKriteria.Text & "'"
        Case Is = 4
            .Find "TanggalLahir = '" & textKriteria.Text & "'"
        Case Is = 5
            .Find "BulanLahir = '" & textKriteria.Text & "'"
        End Select
        If .EOF Then
            MsgBox ("Data tidak ditemukan"), vbExclamation + vbOKOnly, ""
        Else
            Set FormAnggota.DataGrid1.DataSource = FormAnggota.Adodc1.Recordset
            cmBatal.Caption = "&Tutup"
            If FormPengaturan.cekTutupFormCari.Value = Checked Then Me.Hide
        End If
    End With
End If
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub

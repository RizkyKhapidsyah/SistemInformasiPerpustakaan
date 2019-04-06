VERSION 5.00
Begin VB.Form FormTambahJenisPerusahaan 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tambah Jenis"
   ClientHeight    =   855
   ClientLeft      =   4305
   ClientTop       =   2550
   ClientWidth     =   3945
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormTambahJenisPerusahaan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   3945
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.CommandButton cmSimpan 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Simpan"
         Height          =   375
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox textTambahJenisPerusahaan 
         Height          =   390
         Left            =   1200
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tambah Jenis "
         Height          =   270
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   945
      End
   End
End
Attribute VB_Name = "FormTambahJenisPerusahaan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmSimpan_Click()
If textTambahJenisPerusahaan.Text = "" Then
    MsgBox "Silahkan isi jenis perusahaan/lembaga yang baru.", vbExclamation + vbOKOnly, ""
    textTambahJenisPerusahaan.SetFocus
Else
    X = MsgBox("Anda yakin ingin menambahkan '" & textTambahJenisPerusahaan.Text & "' ke dalam daftar jenis perusahaan/lembaga?", vbQuestion + vbYesNo, "Konfirmasi")
    If X = vbYes Then
        With FormBukuMasuk
            .AdodcUntukCMBJenis.Recordset.AddNew
            .AdodcUntukCMBJenis.Recordset.Fields(0).Value = textTambahJenisPerusahaan.Text
            .AdodcUntukCMBJenis.Recordset.Update
            .Refresh
            .MasukkanDatabaseKeAdodcUntukCMBJenis
            .cmbJenis.Text = textTambahJenisPerusahaan.Text
        End With
        Unload Me
    End If
End If
End Sub

Private Sub Form_Load()
    textTambahJenisPerusahaan.Text = ""
End Sub

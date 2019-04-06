VERSION 5.00
Begin VB.Form formTambahKategoriBuku 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tambah Kategori Buku"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4065
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "formTambahKategoriBuku.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4065
   Begin VB.CommandButton cmBatal 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Batal"
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmSimpan 
      BackColor       =   &H00E0E0E0&
      Caption         =   "OK/Save"
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox textKategori 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   1680
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox textID 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   1680
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "cth:123"
      Height          =   270
      Left            =   3120
      TabIndex        =   8
      Top             =   120
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1560
      TabIndex        =   4
      Top             =   600
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Kategori"
      Height          =   270
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   120
   End
End
Attribute VB_Name = "formTambahKategoriBuku"
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
End Sub

Private Sub cmBatal_Click()
    Unload Me
End Sub

Private Sub cmSimpan_Click()
    If textID.Text = "" Then
        MsgBox "Silahkan isi ID untuk kategori yang ingin Anda masukkan", vbExclamation + vbOKOnly, ""
        textID.SetFocus
    ElseIf textKategori.Text = "" Then
        MsgBox "silahkan isi Kategori!", vbExclamation + vbOKOnly, ""
        textKategori.SetFocus
    Else
        X = MsgBox("Apakah Anda yakin ingin menambahkan kategori baru dengan nama : '" & textKategori.Text & "' ?", vbQuestion + vbYesNo, "Konfirmasi")
        If X = vbYes Then
            With FormBukuMasuk
                .AdodcUntukCMBKategori.Recordset.AddNew
                .AdodcUntukCMBKategori.Recordset.Fields(0).Value = textID.Text
                .AdodcUntukCMBKategori.Recordset.Fields(1).Value = textKategori.Text
                .AdodcUntukCMBKategori.Recordset.Update
                .AdodcUntukCMBKategori.Refresh
                .MasukkanDatabaseKeAdodcUntukCMBKategori
                .cmbKategori.Text = textKategori.Text
            End With
            cmBatal.Caption = "&Tutup"
            AturKontrol
            textID.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub


VERSION 5.00
Begin VB.Form FormSorotBukuMasuk 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sorot Data - Buku Masuk"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormSorotBukuMasuk.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   6255
   Begin VB.CommandButton cmBatal 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Batal"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.CommandButton cmSorot 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Sorot"
         Height          =   375
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cmbMode 
         Height          =   390
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   2295
      End
      Begin VB.ComboBox cmbSorotDataBerdasarkan 
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
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FormSorotBukuMasuk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    With cmbSorotDataBerdasarkan
        .Clear
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(0).Name, 0
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(1).Name, 1
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(2).Name, 2
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(3).Name, 3
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(4).Name, 4
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(5).Name, 5
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(6).Name, 6
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(7).Name, 7
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(8).Name, 8
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(9).Name, 9
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(10).Name, 10
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(11).Name, 11
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(12).Name, 12
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(13).Name, 13
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(14).Name, 14
        .AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(15).Name, 15
        .ListIndex = 6
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
    With FormManageBukuMasuk
        .AdodcUtama.Refresh
        .AdodcUtama.RecordSource = "Select * from tbBuku order by " & cmbSorotDataBerdasarkan.Text & " " & cmbMode.Text & ";"
        .AdodcUtama.Refresh
    End With
    cmBatal.Caption = "&Tutup"
    If FormPengaturan.cekTutupFormSorot.Value = Checked Then Me.Hide
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub



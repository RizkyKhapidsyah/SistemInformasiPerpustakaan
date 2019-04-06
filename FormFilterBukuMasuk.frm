VERSION 5.00
Begin VB.Form FormFilterBukuMasuk 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filter Data - Buku Masuk"
   ClientHeight    =   1935
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
   Icon            =   "FormFilterBukuMasuk.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   6255
   Begin VB.CommandButton cmBatal 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Batal"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.CommandButton cmFilter 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Filter"
         Height          =   375
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox cmbMode 
         Height          =   390
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   840
         Width           =   2535
      End
      Begin VB.ComboBox cmbFilterBerdasarkan 
         Height          =   390
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dengan Mode :"
         Height          =   270
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filter Berdasarkan :"
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1290
      End
   End
End
Attribute VB_Name = "FormFilterBukuMasuk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
With Me
    .cmbFilterBerdasarkan.Clear
    .cmbFilterBerdasarkan.AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(0).Name, 0
    .cmbFilterBerdasarkan.AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(1).Name, 1
    .cmbFilterBerdasarkan.AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(2).Name, 2
    .cmbFilterBerdasarkan.AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(3).Name, 3
    .cmbFilterBerdasarkan.AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(4).Name, 4
    .cmbFilterBerdasarkan.AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(5).Name, 5
    .cmbFilterBerdasarkan.AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(6).Name, 6
    .cmbFilterBerdasarkan.AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(7).Name, 7
    .cmbFilterBerdasarkan.AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(8).Name, 8
    .cmbFilterBerdasarkan.AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(9).Name, 9
    .cmbFilterBerdasarkan.AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(10).Name, 10
    .cmbFilterBerdasarkan.AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(11).Name, 11
    .cmbFilterBerdasarkan.AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(12).Name, 12
    .cmbFilterBerdasarkan.AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(13).Name, 13
    .cmbFilterBerdasarkan.AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(14).Name, 14
    .cmbFilterBerdasarkan.AddItem FormManageBukuMasuk.AdodcUtama.Recordset.Fields(15).Name, 15
    .cmbFilterBerdasarkan.ListIndex = 6
    .cmbMode.Clear
    .cmbMode.AddItem "Asc", 0
    .cmbMode.AddItem "Desc", 1
    .cmbMode.ListIndex = 0
End With
End Sub

Private Sub cmBatal_Click()
    Unload Me
End Sub

Private Sub cmFilter_Click()

    With FormManageBukuMasuk
        .AdodcUtama.Refresh
        .AdodcUtama.RecordSource = "Select " & cmbFilterBerdasarkan.Text & " from tbBuku order by " & cmbFilterBerdasarkan.Text & " " & cmbMode.Text & ";"
        .AdodcUtama.Refresh
    End With
    cmBatal.Caption = "&Tutup"
    If FormPengaturan.cekTutupFormFilter.Value = Checked Then Me.Hide
    With FormManageBukuMasuk
        .cmEdit.Enabled = False
        .cmCari.Enabled = False
        .cmSorot.Enabled = False
        .cmFilter.Enabled = False
        .cmHapus.Enabled = False
        .cmBaru.Enabled = False
    End With
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub



VERSION 5.00
Begin VB.Form FormFilterAnggota 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filter Data - Anggota"
   ClientHeight    =   1815
   ClientLeft      =   6270
   ClientTop       =   4920
   ClientWidth     =   4230
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormFilterAnggota.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4230
   Begin VB.ComboBox cmbFilterBerdasarkan 
      Height          =   390
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   3975
   End
   Begin VB.ComboBox cmbMode 
      Height          =   390
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton cmOK 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&OK"
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmBatal 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Batal"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filter Berdasarkan :"
      Height          =   270
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1290
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dengan Mode :"
      Height          =   270
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   915
   End
End
Attribute VB_Name = "FormFilterAnggota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
With Me
    .Caption = "Filter Data - Anggota"
    .cmbFilterBerdasarkan.Clear
    .cmbFilterBerdasarkan.AddItem FormAnggota.Adodc1.Recordset.Fields(0).Name, 0
    .cmbFilterBerdasarkan.AddItem FormAnggota.Adodc1.Recordset.Fields(1).Name, 1
    .cmbFilterBerdasarkan.AddItem FormAnggota.Adodc1.Recordset.Fields(2).Name, 2
    .cmbFilterBerdasarkan.AddItem FormAnggota.Adodc1.Recordset.Fields(3).Name, 3
    .cmbFilterBerdasarkan.AddItem FormAnggota.Adodc1.Recordset.Fields(4).Name, 4
    .cmbFilterBerdasarkan.AddItem FormAnggota.Adodc1.Recordset.Fields(5).Name, 5
    .cmbFilterBerdasarkan.ListIndex = 0
    .cmbMode.Clear
    .cmbMode.AddItem "Asc", 0
    .cmbMode.AddItem "Desc", 1
    .cmbMode.ListIndex = 0
End With
End Sub

Private Sub cmBatal_Click()
    Unload Me
End Sub

Private Sub cmOK_Click()
    With FormAnggota
        .Adodc1.Refresh
        .Adodc1.RecordSource = "Select " & cmbFilterBerdasarkan.Text & " from tbAnggota order by " & cmbFilterBerdasarkan.Text & " " & cmbMode.Text & ";"
        .Adodc1.Refresh
    End With

    cmBatal.Caption = "&Tutup"
    If FormPengaturan.cekTutupFormFilter.Value = Checked Then Me.Hide
    With FormAnggota
        .cmEdit.Enabled = False
        .cmCari.Enabled = False
        .cmSorot.Enabled = False
        .cmFilter.Enabled = False
        .cmHapus.Enabled = False
    End With
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub

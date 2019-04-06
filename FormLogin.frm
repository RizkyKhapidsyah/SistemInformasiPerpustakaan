VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.ocx"
Begin VB.Form FormLogin 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log In"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4410
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4410
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4215
      Begin VB.TextBox textNamaPengguna 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   225
         Width           =   2295
      End
      Begin VB.TextBox textPassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   630
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pengguna"
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   8
         Top             =   225
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1680
         TabIndex        =   7
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   630
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1680
         TabIndex        =   5
         Top             =   630
         Width           =   45
      End
   End
   Begin VB.CommandButton cmDaftar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Daftar"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1215
      Width           =   1095
   End
   Begin VB.CommandButton cmLogin 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Masuk"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5520
      Top             =   3840
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
End
Attribute VB_Name = "FormLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub SambungkanAdodc()
NyambunggUtama
    With Adodc1
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * From tbLogin"
        .Refresh
    End With
End Sub
Sub AturKontrol()
    For Each Objek In Me
        If TypeName(Objek) = "TextBox" Then
            Objek.Text = ""
        End If
    Next
End Sub

Private Sub cmDaftar_Click()
    With FormDaftarPenggunaBaru
        .TextPasswordLama.Enabled = False
        .TextPasswordLama.BackColor = FormDaftarPenggunaBaru.BackColor
        .Show vbModal, Me
    End With
End Sub

Private Sub cmLogin_Click()
    If RS.State = 1 Then RS.Close
    RS.Open "select * from tbLogin where NamaPengguna = '" & textNamaPengguna.Text & "' And Password = '" & textPassword.Text & "' And Password = '" & textPassword.Text & "'", CN, 3, 3
        If textNamaPengguna.Text = "" Then
            MsgBox "Silahkan Isi Nama Pengguna Anda!", vbExclamation + vbOKOnly, "MainSystem : Login"
            textNamaPengguna.SetFocus
        ElseIf textPassword.Text = "" Then
            MsgBox "Silahkan Isi Password Anda!", vbExclamation + vbOKOnly, "MainSystem : Login"
            textPassword.SetFocus
        Else
            If Not RS.EOF Then
                With FORM_UTAMA
                    .StatusBawah.Panels.Item(2).Text = textNamaPengguna.Text
                    .menuAdmin.Enabled = True
                    .menuTransaksi.Enabled = True
                    .menuLaporan.Enabled = True
                    .Show
                End With
                Unload Me
                    FormSplash.Show vbModal, Me
            Else
                MsgBox "Maaf, Data login yang Anda masukkan Tidak sesuai dengan data Login!" & vbCrLf & _
                        "Silahkan coba lagi!", vbCritical + vbOKOnly, ""
                textNamaPengguna.SetFocus
            End If
        End If
End Sub

Private Sub Form_Load()
    AturKontrol
    SambungkanAdodc
    If Adodc1.Recordset.RecordCount = 0 Then
        With FormDaftarPenggunaBaru
            .Show vbModal, Me
        End With
        Unload Me
    End If
End Sub

Private Sub textNamaPengguna_Change()
    If textNamaPengguna.Text = "" Then
        If textPassword.Text = "" Then
            cmLogin.Enabled = False
        Else
            cmLogin.Enabled = True
        End If
    Else
        cmLogin.Enabled = True
    End If
End Sub

Private Sub textPassword_Change()
    If textPassword.Text = "" Then
        If textNamaPengguna.Text = "" Then
            cmLogin.Enabled = False
        Else
            cmLogin.Enabled = True
        End If
    Else
        cmLogin.Enabled = True
    End If
End Sub

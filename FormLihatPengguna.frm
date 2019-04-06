VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormLihatPengguna 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daftar Pengguna"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12105
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormLihatPengguna.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   12105
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Table View"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   33
      Top             =   120
      Width           =   5655
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4575
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   8070
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Paralel View"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   5880
      TabIndex        =   5
      Top             =   120
      Width           =   6135
      Begin VB.TextBox textTahun 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   5040
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox textBulan 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   4440
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox textTanggal 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   3840
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox textTempat 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   2640
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox TextJenisKelamin 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   2640
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox textNamaPenggunaLogin 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   2640
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox textNamaAsli 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   2640
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox textAlamat 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   960
         Left            =   2640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Text            =   "FormLihatPengguna.frx":000C
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox textStatusPendidikan 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   2640
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   2880
         Width           =   3375
      End
      Begin VB.TextBox textStatusPekerjaan 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   2640
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   3360
         Width           =   3375
      End
      Begin VB.TextBox textStatusHubungan 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   2640
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   3840
         Width           =   3375
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tempat Tanggal Lahir"
         Height          =   240
         Left            =   240
         TabIndex        =   32
         Top             =   3840
         Width           =   1320
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   240
         Left            =   2520
         TabIndex        =   31
         Top             =   3840
         Width           =   45
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Pekerjaan"
         Height          =   240
         Left            =   240
         TabIndex        =   30
         Top             =   3360
         Width           =   1065
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   240
         Left            =   2520
         TabIndex        =   29
         Top             =   3360
         Width           =   45
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Pendidikan"
         Height          =   240
         Left            =   240
         TabIndex        =   28
         Top             =   2880
         Width           =   1125
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   240
         Left            =   2520
         TabIndex        =   27
         Top             =   2880
         Width           =   45
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         Height          =   240
         Left            =   240
         TabIndex        =   26
         Top             =   1800
         Width           =   435
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   240
         Left            =   2520
         TabIndex        =   25
         Top             =   1800
         Width           =   45
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tempat Tanggal Lahir"
         Height          =   240
         Left            =   240
         TabIndex        =   20
         Top             =   1440
         Width           =   1320
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   240
         Left            =   2520
         TabIndex        =   19
         Top             =   1440
         Width           =   45
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Kelamin"
         Height          =   240
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   240
         Left            =   2520
         TabIndex        =   16
         Top             =   1080
         Width           =   45
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Asli"
         Height          =   240
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   240
         Left            =   2520
         TabIndex        =   14
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   240
         Left            =   2520
         TabIndex        =   7
         Top             =   360
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pengguna [Login]"
         Height          =   240
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1530
      End
   End
   Begin VB.CommandButton cmAkhir 
      BackColor       =   &H00E0E0E0&
      Caption         =   ">>"
      Height          =   375
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton cmSelanjutnya 
      BackColor       =   &H00E0E0E0&
      Caption         =   ">"
      Height          =   375
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox textNamaPenggunaData 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   4560
      Width           =   2280
   End
   Begin VB.CommandButton cmSebelumnya 
      BackColor       =   &H00E0E0E0&
      Caption         =   "<"
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   495
   End
   Begin MSAdodcLib.Adodc AdodcDaftarPengguna_Lihat 
      Height          =   330
      Left            =   0
      Top             =   0
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
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmAwal 
      BackColor       =   &H00E0E0E0&
      Caption         =   "<<"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4560
      Width           =   495
   End
End
Attribute VB_Name = "FormLihatPengguna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub SambungkanAdodc()
    NyambunggUtama
    With AdodcDaftarPengguna_Lihat
        .ConnectionString = CN.ConnectionString
        .RecordSource = "select NamaPengguna,NamaAsli,JenisKelamin,TempatLahir,TanggalLahir,BulanLahir,TahunLahir,Alamat,StatusPendidikan,StatusPekerjaan,StatusHubungan From tbLogin"
        Set DataGrid1.DataSource = AdodcDaftarPengguna_Lihat
        .Refresh
    End With
End Sub
Sub AturKontrol()
    For Each Objek In Me
        If TypeName(Objek) = "TextBox" Then
            With Objek
                .Text = ""
                .MaxLength = 254
                .Locked = True
            End With
        End If
    Next
    With Me
        .DataGrid1.AllowUpdate = False
    End With
    AlirkanDataKeOutput
End Sub
Sub AlirkanDataKeOutput()
On Error Resume Next
    With Me
        .textNamaPenggunaLogin.Text = AdodcDaftarPengguna_Lihat.Recordset(0).Value
        .textNamaAsli.Text = AdodcDaftarPengguna_Lihat.Recordset(1).Value
        .TextJenisKelamin.Text = AdodcDaftarPengguna_Lihat.Recordset(2).Value
        .textTempat.Text = AdodcDaftarPengguna_Lihat.Recordset(3).Value
        .textTanggal.Text = AdodcDaftarPengguna_Lihat.Recordset(4).Value
        .textBulan.Text = AdodcDaftarPengguna_Lihat.Recordset(5).Value
        .textTahun.Text = AdodcDaftarPengguna_Lihat.Recordset(6).Value
        .textAlamat.Text = AdodcDaftarPengguna_Lihat.Recordset(7).Value
        .textStatusPendidikan.Text = AdodcDaftarPengguna_Lihat.Recordset(8).Value
        .textStatusPekerjaan.Text = AdodcDaftarPengguna_Lihat.Recordset(9).Value
        .textStatusHubungan.Text = AdodcDaftarPengguna_Lihat.Recordset(10).Value
        .textNamaPenggunaData.Text = AdodcDaftarPengguna_Lihat.Recordset(0).Value
    End With
End Sub

Private Sub cmAkhir_Click()
    AdodcDaftarPengguna_Lihat.Recordset.MoveLast
    AlirkanDataKeOutput
End Sub

Private Sub cmAwal_Click()
    AdodcDaftarPengguna_Lihat.Recordset.MoveFirst
    AlirkanDataKeOutput
End Sub


Private Sub cmSebelumnya_Click()
    If AdodcDaftarPengguna_Lihat.Recordset.BOF = True Then
        AdodcDaftarPengguna_Lihat.Recordset.MoveLast
        AlirkanDataKeOutput
    End If
        AdodcDaftarPengguna_Lihat.Recordset.MovePrevious
        AlirkanDataKeOutput
End Sub

Private Sub cmSelanjutnya_Click()
    If AdodcDaftarPengguna_Lihat.Recordset.EOF = True Then
        AdodcDaftarPengguna_Lihat.Recordset.MoveFirst
        AlirkanDataKeOutput
    End If
        AdodcDaftarPengguna_Lihat.Recordset.MoveNext
        AlirkanDataKeOutput
End Sub

Private Sub DataGrid1_Click()
    DataGrid1_DblClick
End Sub

Private Sub DataGrid1_DblClick()
    AlirkanDataKeOutput
End Sub

Private Sub Form_Load()
    SambungkanAdodc
    AturKontrol
End Sub

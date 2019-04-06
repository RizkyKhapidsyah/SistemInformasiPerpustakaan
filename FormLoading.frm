VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FormLoading 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sedang Memuat..."
   ClientHeight    =   510
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4275
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormLoading.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   510
   ScaleWidth      =   4275
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   120
      Top             =   2400
   End
   Begin MSComctlLib.ProgressBar Proses 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.Label LabelPersen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "---"
      Height          =   270
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   180
   End
End
Attribute VB_Name = "FormLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Proses.Value = 1
End Sub

Private Sub Timer1_Timer()
    Proses.Value = Proses.Value + 1
    LabelPersen.Caption = Proses.Value & "%"
    If Proses.Value = 100 Then
        Timer1.Enabled = False
        FormLogin.Show
        Unload Me
    End If
End Sub

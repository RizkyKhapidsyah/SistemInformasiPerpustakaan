VERSION 5.00
Begin VB.Form FormPengaturan 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pengaturan"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3960
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormPengaturan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Data"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   3735
      Begin VB.CheckBox cekTutupFormFilter 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tutup Form Filter setelah data difilter"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   3375
      End
      Begin VB.CheckBox cekTutupFormSorot 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tutup Form Sorot setelah data disortir"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   3375
      End
      Begin VB.CheckBox cekTutupFormCari 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tutup Form Cari setelah data ditemukan"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame FrameTabel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tabel"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.CheckBox cekAutoRefresh 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Auto Refresh"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox cekKunciTabel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Kunci Tabel"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   1  'Checked
         Width           =   2415
      End
   End
End
Attribute VB_Name = "FormPengaturan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


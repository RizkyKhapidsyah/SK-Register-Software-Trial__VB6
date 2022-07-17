VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registering Example"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   3495
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer 
      Interval        =   50
      Left            =   120
      Top             =   720
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.PictureBox picReg 
      AutoRedraw      =   -1  'True
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.Label lblRegCode 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblEmail 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblCompany 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblUser 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblRegName 
         BackStyle       =   0  'Transparent
         Caption         =   "Registered User:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    DoGradient Me
    DoGradient picReg
    lblUser.Caption = GetSetting("Program", "Info", "User")
    lblCompany.Caption = GetSetting("Program", "Info", "Company")
    lblEmail.Caption = GetSetting("Program", "Info", "Email")
    lblRegCode.Caption = GetSetting("Program", "Info", "RegID")
End Sub


Private Sub cmdRegister_Click()
    Dim DateRan
    DateRan = Format(Date, "dd-mmmm-yyyy")
    SaveSetting "Program", "Info", "RunOnce", "0"
    SaveSetting "Program", "Info", "FirstRan", DateRan
    SaveSetting "Program", "Info", "LastRan", DateRan
    Unload Me
    Call Main
End Sub

Private Sub cmdClear_Click()
    Dim DateRan
    DateRan = Format(Date, "dd-mmmm-yyyy")
    DeleteSetting "Program", "Info"
    SaveSetting "Program", "Info", "RunOnce", "0"
    SaveSetting "Program", "Info", "FirstRan", DateRan
    SaveSetting "Program", "Info", "LastRan", DateRan
    Form_Load
End Sub

Private Sub cmdExit_Click()
    Unload Me
    End
End Sub

Private Sub Timer_Timer()
'    lblRegName.Top = lblRegName.Top - 1
'    lblUser.Top = lblUser.Top - 1
'    lblCompany.Top = lblCompany.Top - 1
'    lblEmail.Top = lblEmail.Top - 1
'    lblRegCode.Top = lblRegCode.Top - 1
End Sub


VERSION 5.00
Begin VB.Form frmRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register The Software"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRegCode 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin VB.TextBox txtCompany 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4335
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   4335
   End
   Begin VB.CommandButton cmdReg 
      Caption         =   "Register"
      Default         =   -1  'True
      Height          =   615
      Left            =   3240
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "User's Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label lblCompany 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "User's Company"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Label lblEmail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "User's Email Address"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Label lblRegCode 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Registration Number"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   2895
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    DoGradient Me
    txtName.Text = GetSetting("Program", "Info", "User", "Your Name")
    txtCompany.Text = GetSetting("Program", "Info", "Company", "Your Company")
    txtEmail.Text = GetSetting("Program", "Info", "Email", "your@email.com")
    txtRegCode.Text = GetSetting("Program", "Info", "RegID")
End Sub

Private Sub cmdReg_Click()
    SaveSetting "Program", "Info", "User", txtName.Text
    SaveSetting "Program", "Info", "Company", txtCompany.Text
    SaveSetting "Program", "Info", "Email", txtEmail.Text
    SaveSetting "Program", "Info", "RegID", txtRegCode.Text
    SaveSetting "Program", "Info", "RunOnce", "1"
    If RegTest = False Then
        MsgBox "Registration Code Invalid!" & vbCrLf & "Please enter the correct code.", vbCritical + vbOKOnly, "Invalid Registration Code"
        txtRegCode.SetFocus
    End If
    Unload Me
    Call Main
End Sub

Private Sub txtName_GotFocus()
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
End Sub

Private Sub txtCompany_GotFocus()
    txtCompany.SelStart = 0
    txtCompany.SelLength = Len(txtCompany.Text)
End Sub

Private Sub txtEmail_GotFocus()
    txtEmail.SelStart = 0
    txtEmail.SelLength = Len(txtEmail.Text)
End Sub

Private Sub txtRegCode_GotFocus()
    txtRegCode.SelStart = 0
    txtRegCode.SelLength = Len(txtRegCode.Text)
End Sub


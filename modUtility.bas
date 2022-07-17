Attribute VB_Name = "modUtility"
Option Explicit
        
Sub Main()
    Dim RunOnce As String, DateRan As String, ExpireDate As String
    Dim Date1 As Date, Date2 As Date, ExpDateDiff As Integer
    Dim Register As Integer, RegID As Variant
    RunOnce = GetSetting("Program", "Info", "RunOnce")
    DateRan = Format(Date, "dd-mmmm-yyyy")
    Select Case RunOnce
        Case 0      ' This is run when the program is unregistered.
            ExpireDate = GetSetting("Program", "Info", "FirstRan")
            Date1 = Format(ExpireDate, "m/d/yyyy")
            Date2 = Format(DateRan, "m/d/yyyy")
            ExpDateDiff = DateDiff("d", Date1, Date2)
            If ExpDateDiff > 30 Then
                Register = MsgBox("This program may be used for 30 days." & vbCrLf & "Would you like to register to continue use.", vbCritical + vbYesNo, "Please register")
                Select Case Register
                    Case 6
                        frmRegister.Show
                    Case 7
                        End
                End Select
            Else
                Register = MsgBox("This program may be used for 30 days." & vbCrLf & "Would you like to register to continue use.", vbCritical + vbYesNo, "Please register")
                Select Case Register
                    Case 6
                        frmRegister.Show
                    Case 7
                        frmMain.Show
                        SaveSetting "Program", "Info", "LastRan", DateRan
                End Select
            End If
        Case 1      ' This is run when the program has been registered.
            If RegTest = True Then
                frmMain.Show
            Else
                frmRegister.Show
            End If
        Case Else   ' This is run when the registry setting is not found
            SaveSetting "Program", "Info", "RunOnce", "0"
            SaveSetting "Program", "Info", "FirstRan", DateRan
            SaveSetting "Program", "Info", "LastRan", DateRan
            MsgBox "This program may be used for 30 days." & vbCrLf & "You will be prompted to register next" & vbCrLf & "time you run the program.", vbCritical + vbOKOnly, "Please register"
            Call Main
    End Select
End Sub

Function RegTest() As Boolean
    Dim RegNum As String
    RegNum = GetSetting("Program", "Info", "RegID")
    If RegNum = "123456789" Then
        RegTest = True
    Else
        RegTest = False
    End If
End Function

Public Sub DoGradient(FormName As Object)
On Error Resume Next
    Dim i As Integer, y As Integer
    FormName.AutoRedraw = True
    FormName.DrawStyle = 6
    FormName.DrawMode = 13
    FormName.DrawWidth = 13
    FormName.ScaleMode = 3
    FormName.ScaleHeight = 256
    For i = 0 To 510
        FormName.Line (0, y)-(FormName.Width, y + 1), RGB(0, 0, i), BF
        y = y + 1
    Next i
End Sub


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "驾考科三灯光模拟"
   ClientHeight    =   3435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5760
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Reference to https://stackoverflow.com/a/38570484

Option Explicit

Dim sButton As String

Private Sub ToggleButton11_Click()

    ButtonLoad 11

End Sub

Private Sub ToggleButton12_Click()

    ButtonLoad 12

End Sub

Private Sub ToggleButton13_Click()

    ButtonLoad 13

End Sub

Private Sub ToggleButton21_Click()

    ButtonLoad 21

End Sub

Private Sub ToggleButton22_Click()

    ButtonLoad 22

End Sub

Private Sub ToggleButton23_Click()

    ButtonLoad 23

End Sub

Private Sub ToggleButton24_Click()

    ButtonLoad 24

End Sub

Private Sub ToggleButton25_Click()

    ButtonLoad 25

End Sub

Private Sub ToggleButton31_Click()

    ButtonLoad 31

End Sub

Sub ButtonLoad(iButton As Integer)

    Select Case iButton
        Case 11
        '关闭
            If sButton = "" Then
                sButton = "11"
                    Me.ToggleButton11.Value = True
                    Me.ToggleButton12.Value = False
                    Me.ToggleButton13.Value = False
                    LightRecord = LightRecord & LightStatus
                sButton = ""
            End If
        Case 12
        '示廓
            If sButton = "" Then
                sButton = "12"
                    Me.ToggleButton11.Value = False
                    Me.ToggleButton12.Value = True
                    Me.ToggleButton13.Value = False
                    LightRecord = LightRecord & LightStatus
                sButton = ""
            End If
        Case 13
        '近光
            If sButton = "" Then
                sButton = "13"
                    Me.ToggleButton11.Value = False
                    Me.ToggleButton12.Value = False
                    Me.ToggleButton13.Value = True
                    LightRecord = LightRecord & LightStatus
                sButton = ""
            End If
        Case 21
        '关闭
            If sButton = "" Then
                sButton = "21"
                    Me.ToggleButton21.Value = True
                    Me.ToggleButton22.Value = False
                    Me.ToggleButton23.Value = False
                    Me.ToggleButton24.Value = False
                    Me.ToggleButton25.Value = False
                    LightRecord = LightRecord & LightStatus
                sButton = ""
            End If
        Case 22
        '闪光
            If sButton = "" Then
                sButton = "22"
                    Me.ToggleButton21.Value = False
                    Me.ToggleButton22.Value = True
                    Me.ToggleButton23.Value = False
                    Me.ToggleButton24.Value = False
                    Me.ToggleButton25.Value = False
                    LightRecord = LightRecord & LightStatus
                    Me.ToggleButton21.Value = True
                    Me.ToggleButton22.Value = False
                    Me.ToggleButton23.Value = False
                    Me.ToggleButton24.Value = False
                    Me.ToggleButton25.Value = False
                sButton = ""
            End If
        Case 23
        '远光
            If sButton = "" Then
                sButton = "23"
                    Me.ToggleButton21.Value = False
                    Me.ToggleButton22.Value = False
                    Me.ToggleButton23.Value = True
                    Me.ToggleButton24.Value = False
                    Me.ToggleButton25.Value = False
                    LightRecord = LightRecord & LightStatus
                sButton = ""
            End If
        Case 24
        '左转
            If sButton = "" Then
                sButton = "24"
                    Me.ToggleButton21.Value = False
                    Me.ToggleButton22.Value = False
                    Me.ToggleButton23.Value = False
                    Me.ToggleButton24.Value = True
                    Me.ToggleButton25.Value = False
                    LightRecord = LightRecord & LightStatus
                sButton = ""
            End If
        Case 25
        '右转
            If sButton = "" Then
                sButton = "25"
                    Me.ToggleButton21.Value = False
                    Me.ToggleButton22.Value = False
                    Me.ToggleButton23.Value = False
                    Me.ToggleButton24.Value = False
                    Me.ToggleButton25.Value = True
                    LightRecord = LightRecord & LightStatus
                sButton = ""
            End If
        Case 31
        '双跳
            If sButton = "" Then
                sButton = "31"
                    LightRecord = LightRecord & LightStatus
                sButton = ""
            End If
    End Select

End Sub

Public Function LightStatus() As String

    If Me.ToggleButton11.Value = True Then LightStatus = "11"
    If Me.ToggleButton12.Value = True Then LightStatus = "12"
    If Me.ToggleButton13.Value = True Then LightStatus = "13"
    
    If Me.ToggleButton21.Value = True Then LightStatus = LightStatus & "21"
    If Me.ToggleButton22.Value = True Then LightStatus = LightStatus & "22"
    If Me.ToggleButton23.Value = True Then LightStatus = LightStatus & "23"
    If Me.ToggleButton24.Value = True Then LightStatus = LightStatus & "24"
    If Me.ToggleButton25.Value = True Then LightStatus = LightStatus & "25"
    
    If Me.ToggleButton31.Value = True Then LightStatus = LightStatus & "31" Else LightStatus = LightStatus & "30"

End Function

Private Sub UserForm_Click()
    Dim n As Byte

    'Fill empty LightRecord by current LightStatus if light unchanged
    If "" = LightRecord Then
        LightRecord = LightStatus
    End If
    
   'MsgBox LightRecord  ' Show LightRecord for debug
    LightCheck
    
    'Next light
    Randomize
    If 12 = LightNumber Then
        n = 1
    ElseIf 1 = LightNumber Then
        n = 2
    Else
        Do
            n = Int((12 - 3 + 1) * Rnd + 3)
        Loop While n = LightNumber
    End If
    
    LightNumber = n
    Me.Caption = Sheet1.Cells(1, LightNumber)
    LightRecord = ""

End Sub

Private Sub UserForm_Initialize()

    sButton = "11"
        Me.ToggleButton11.Value = True
        Me.ToggleButton12.Value = False
        Me.ToggleButton13.Value = False
    sButton = ""
    
    sButton = "21"
        Me.ToggleButton21.Value = True
        Me.ToggleButton22.Value = False
        Me.ToggleButton23.Value = False
        Me.ToggleButton24.Value = False
        Me.ToggleButton25.Value = False
    sButton = ""
    
    sButton = "31"
        Me.ToggleButton31.Value = False
    sButton = ""

    LightNumber = 1
    Me.Caption = Sheet1.Cells(1, LightNumber)
    LightRecord = ""
    
End Sub

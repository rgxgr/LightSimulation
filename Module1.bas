Attribute VB_Name = "Module1"
Option Explicit

Public LightNumber As Byte

Public LightRecord As String


Sub LightCheck()
    Dim Result As String
    
    With Sheet1
        Select Case LightNumber
            Case Is = 1   '请检查灯光
                Result = "132230132230"
            Case Is = 2   '开启灯光
                Result = "132130"
            Case Is = 3   '超车
                Result = "132430132130132230132230132530132130"
            Case Is = 4   '右转弯通过路口
                Result = "132530"
            Case Is = 5   '直线通过路口
                Result = "132130"
            Case Is = 6   '通过人行橫道线
                Result = "132230132230"
            Case Is = 7   '会车
                Result = "132130"
            Case Is = 8   '路边临时停车
                Result = "122131"
            Case Is = 9   '照明不良道路行驶
                Result = "132330"
            Case Is = 10  '左转弯通过路口
                Result = "132430"
            Case Is = 11  '通过桥梁、急弯、坡道行驶
                Result = "132230132230"
            Case Is = 12  '模拟夜间考试完成，请关闭所有灯光
                Result = "112130"
        End Select
        
        'Fill Result after last unempty cell
        .Cells(.Cells(1048576, LightNumber).End(xlUp).row + 1, LightNumber) = Abs(Result = Right(LightRecord, Len(Result))) * 2 - 1
    End With
End Sub

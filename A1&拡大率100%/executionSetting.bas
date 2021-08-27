Attribute VB_Name = "executionSetting"
'/**
' * 処理実行前の設定
' * 基本的な目的は処理の高速化
'**/
Sub beforeExecution()
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
End Sub


'/**
' * 処理実行後は処理実行前の状態に戻す
'**/
Sub afterExecution()
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub


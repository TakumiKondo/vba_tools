Attribute VB_Name = "executionSetting"
'/**
' * �������s�O�̐ݒ�
' * ��{�I�ȖړI�͏����̍�����
'**/
Sub beforeExecution()
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
End Sub


'/**
' * �������s��͏������s�O�̏�Ԃɖ߂�
'**/
Sub afterExecution()
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub


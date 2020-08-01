VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "�o�͐ݒ�"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6375
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' /**
'  * �ꊇ�o��
' /*
Private Sub AllOutputButton_Click()
    ' �e�ݒ�l��Validation
    If settingValidation() = False Then
        Exit Sub
    End If
    
    ' �t�@�C������
    Set targetSheets = getTargetSheets()
    For Each st In targetSheets
        createFile (st)
    Next st
    
    MsgBox completed()
End Sub

' /**
'  * �ʏo��
' /*
Private Sub SelectOutputButton_Click()
    ' �t�@�C���I��K�{�i�����I��ݒ莞�j
    ' �i1�ȏ�I�����Ă��邱�Ɓj
    If SheetList.MultiSelect = fmMultiSelectMulti Then
        For i = 0 To SheetList.ListCount - 1
            If SheetList.Selected(i) = True Then
                GoTo CONTINUE
                Debug.Print i & " : " & "seleted"
            End If
        Next i
        MsgBox NotSelectedList()
        Exit Sub
CONTINUE:
    End If

    ' �e�ݒ�l��Validation
    If settingValidation() = False Then
        Exit Sub
    End If
    
    ' �t�@�C������
    For i = 0 To SheetList.ListCount - 1
        If SheetList.Selected(i) = True Then
            createFile (SheetList.List(i))
        End If
    Next i
    
    MsgBox completed()
End Sub


' /**
'  * �t�@�C������
' /*
Sub createFile(sheetName As String)

    Dim fileName As String: fileName = outputPath() & "\" & sheetName & "." & outputType()
    Set targetSheet = ThisWorkbook.Worksheets(sheetName)
    MaxCol = targetSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    MaxRow = targetSheet.Cells(Rows.Count, 1).End(xlUp).Row

    With CreateObject("ADODB.Stream")
        .Charset = charCode()
        .Open
        For Row = 1 To MaxRow
            For col = 1 To MaxCol
                .WriteText """" & targetSheet.Cells(Row, col) & """"
                If col < MaxCol Then
                    .WriteText delimiter()
                End If
                If col = MaxCol Then
                    .WriteText newLineCode()
                End If
            Next col
        Next Row
    
        ' UTF-8(BOM����)�ł���΁ABOM���폜
        If CharCodeBox.Text = "UTF-8(BOM����)" Then
            '�X�g���[���̈ʒu��0����3�ɂ���
            .Position = 0
            .Type = 1   ' Binary Type �ɕύX
            .Position = 3

            ' �f�[�^�ޔ�
            Dim nonBomData
            nonBomData = .Read
            .Close

            ' �ޔ��f�[�^�̏�������
            .Open
            .Write nonBomData
        End If
        
        ' �ʖ��ۑ�
        .SaveToFile saveFileName(fileName), 1
        .Close
        
    End With
End Sub



' /**
'  * �ۑ��t�@�C�����̐ݒ�
' /*
Function saveFileName(fileName As String) As String
    ' �t�@�C�������݂��Ȃ��ꍇ�͕ύX���Ȃ�
    If Dir(fileName) = "" Then
        saveFileName = fileName
        Exit Function
    End If
    
    ' �t�@�C�������݂���ꍇ�A�^�C���X�^���v��t�^�������O��ݒ肷��
    extention = Right(fileName, 4)
    baseFileName = Left(fileName, Len(fileName) - 4)
    saveFileName = baseFileName & "_" & Format(Now, "yyyymmdd_hhmmss") & extention
End Function


' /**
'  * �e�ݒ�l��Validation
' /*
Private Function settingValidation() As Boolean
    IsValid = True
    
    ' �����R�[�h�K�{
    If CharCodeBox.ListIndex = -1 Then
        MsgBox NotSelectedItem("�����R�[�h")
        IsValid = False
    End If
    
    ' ���s�R�[�h�K�{
    If NewLineCodeBox.ListIndex = -1 Then
        MsgBox NotSelectedItem("���s�R�[�h")
        IsValid = False
    End If
    
    ' �o�͌`���K�{
    If OutputTypeBox.ListIndex = -1 Then
        MsgBox NotSelectedItem("�o�͌`��")
        IsValid = False
    End If
    
    settingValidation = IsValid
End Function


' /**
'  * �����R�[�h�̎擾
' /*
Private Function charCode()
    If CharCodeBox.Text Like "UTF-8*" Then
        charCode = "UTF-8"
        Exit Function
    End If
    If CharCodeBox.Text = "SJIS" Then
        charCode = "Shift-JIS"
        Exit Function
    End If
End Function


' /**
'  * ���s�R�[�h�̎擾
' /*
Private Function newLineCode()
    If NewLineCodeBox.Text = "CRLF" Then
        newLineCode = vbCrLf
        Exit Function
    End If
    If NewLineCodeBox.Text = "CR" Then
        newLineCode = vbCr
        Exit Function
    End If
    If NewLineCodeBox.Text = "LF" Then
        newLineCode = vbLf
        Exit Function
    End If
End Function


' /**
'  * �o�͌`���̎擾
' /*
Private Function outputType() As String
    outputType = OutputTypeBox.Text
End Function


' /**
'  * ��؂蕶���̎擾
' /*
Private Function delimiter() As String
    If OutputTypeBox.Text = "CSV" Then
        delimiter = ","
        Exit Function
    End If
    If OutputTypeBox.Text = "TSV" Then
        delimiter = vbTab
        Exit Function
    End If
End Function


' /**
'  * �t�H�[���̏�����
' /*
Private Sub UserForm_Initialize()
    ' �o�͑ΏۃV�[�g�̎擾
    Set targetSheets = getTargetSheets()
    For Each st In targetSheets
        SheetList.AddItem st
    Next st
    
    ' �����R�[�h�̏�����
    With CharCodeBox
        .AddItem "UTF-8(BOM����)"
        .AddItem "UTF-8(BOM����)"
        .AddItem "SJIS"
        .ListIndex = 0
    End With
    
    ' ���s�R�[�h�̏�����
    With NewLineCodeBox
        .AddItem "CRLF"
        .AddItem "CR"
        .AddItem "LF"
        .ListIndex = 0
    End With
    
    ' �o�͌`���̏�����
    With OutputTypeBox
        .AddItem "CSV"
        .AddItem "TSV"
        .ListIndex = 0
    End With
End Sub

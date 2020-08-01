VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "出力設定"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6375
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' /**
'  * 一括出力
' /*
Private Sub AllOutputButton_Click()
    ' 各設定値のValidation
    If settingValidation() = False Then
        Exit Sub
    End If
    
    ' ファイル生成
    Set targetSheets = getTargetSheets()
    For Each st In targetSheets
        createFile (st)
    Next st
    
    MsgBox completed()
End Sub

' /**
'  * 個別出力
' /*
Private Sub SelectOutputButton_Click()
    ' ファイル選択必須（複数選択設定時）
    ' （1つ以上選択していること）
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

    ' 各設定値のValidation
    If settingValidation() = False Then
        Exit Sub
    End If
    
    ' ファイル生成
    For i = 0 To SheetList.ListCount - 1
        If SheetList.Selected(i) = True Then
            createFile (SheetList.List(i))
        End If
    Next i
    
    MsgBox completed()
End Sub


' /**
'  * ファイル生成
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
    
        ' UTF-8(BOM無し)であれば、BOMを削除
        If CharCodeBox.Text = "UTF-8(BOM無し)" Then
            'ストリームの位置を0から3にする
            .Position = 0
            .Type = 1   ' Binary Type に変更
            .Position = 3

            ' データ退避
            Dim nonBomData
            nonBomData = .Read
            .Close

            ' 退避データの書き込み
            .Open
            .Write nonBomData
        End If
        
        ' 別名保存
        .SaveToFile saveFileName(fileName), 1
        .Close
        
    End With
End Sub



' /**
'  * 保存ファイル名の設定
' /*
Function saveFileName(fileName As String) As String
    ' ファイルが存在しない場合は変更しない
    If Dir(fileName) = "" Then
        saveFileName = fileName
        Exit Function
    End If
    
    ' ファイルが存在する場合、タイムスタンプを付与した名前を設定する
    extention = Right(fileName, 4)
    baseFileName = Left(fileName, Len(fileName) - 4)
    saveFileName = baseFileName & "_" & Format(Now, "yyyymmdd_hhmmss") & extention
End Function


' /**
'  * 各設定値のValidation
' /*
Private Function settingValidation() As Boolean
    IsValid = True
    
    ' 文字コード必須
    If CharCodeBox.ListIndex = -1 Then
        MsgBox NotSelectedItem("文字コード")
        IsValid = False
    End If
    
    ' 改行コード必須
    If NewLineCodeBox.ListIndex = -1 Then
        MsgBox NotSelectedItem("改行コード")
        IsValid = False
    End If
    
    ' 出力形式必須
    If OutputTypeBox.ListIndex = -1 Then
        MsgBox NotSelectedItem("出力形式")
        IsValid = False
    End If
    
    settingValidation = IsValid
End Function


' /**
'  * 文字コードの取得
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
'  * 改行コードの取得
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
'  * 出力形式の取得
' /*
Private Function outputType() As String
    outputType = OutputTypeBox.Text
End Function


' /**
'  * 区切り文字の取得
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
'  * フォームの初期化
' /*
Private Sub UserForm_Initialize()
    ' 出力対象シートの取得
    Set targetSheets = getTargetSheets()
    For Each st In targetSheets
        SheetList.AddItem st
    Next st
    
    ' 文字コードの初期化
    With CharCodeBox
        .AddItem "UTF-8(BOM無し)"
        .AddItem "UTF-8(BOMあり)"
        .AddItem "SJIS"
        .ListIndex = 0
    End With
    
    ' 改行コードの初期化
    With NewLineCodeBox
        .AddItem "CRLF"
        .AddItem "CR"
        .AddItem "LF"
        .ListIndex = 0
    End With
    
    ' 出力形式の初期化
    With OutputTypeBox
        .AddItem "CSV"
        .AddItem "TSV"
        .ListIndex = 0
    End With
End Sub

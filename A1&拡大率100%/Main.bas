Attribute VB_Name = "Main"
Sub Main()
    Dim Path As String: Path = ThisWorkbook.Worksheets("Sheet1").Range("targetDir").Value   '// 対象フォルダパス
    Const cnsTitle = "フォルダ内のエクセルファイル提出状態化" 'ダイアログのタイトル
    
    Dim Workbook As Workbook
    Dim sheet As Object
    
    ' フォルダの存在確認
    If Dir(Path, vbDirectory) = "" Then
        MsgBox "指定のフォルダは存在しません。", vbExclamation, cnsTitle
        Exit Sub
    End If
    
    
    Call beforeExecution
    '// A1＆拡大率100%を設定する
    Call setA1And100Per(Path)
    Call afterExecution
    
    MsgBox "完了しました。"

End Sub

'// A1＆拡大率100%を設定する
Private Sub setA1And100Per(Path)
    '// A1＆拡大率100％を設定する（ルートフォルダ）
    Call executeA1And100Per(Path)
    
    '// サブフォルダを再帰する
    If ActiveSheet.CheckBoxes("CheckBox").Value = 1 Then
        recursion (Path)
    End If

End Sub


'// Excel系ファイルのみ、A1かつ拡大率100％にして、最初のシートを指定して保存する
Private Sub executeA1And100Per(Path)
    Const cnsDIR = "\*.*"
    Dim strFileName As String '処理中のファイル名を格納する変数
    Dim fileAndPath As String '処理中のファイル名（パス含む）を格納する変数
    Dim pos As Long

    ' 先頭のファイル名の取得
    strFileName = Dir(Path & cnsDIR, vbNormal)
    ' ファイルが見つからなくなるまで繰り返す
    Do While strFileName <> ""
    
        ' エクセルファイルのみを処理対象とする
        pos = InStrRev(strFileName, ".")
        If Not LCase(Mid(strFileName, pos + 1)) Like "xls*" Then
            ' 次のファイル名を取得
            GoTo Continue
        End If
        
        ' 自ファイル（A1&拡大率100％.xlsm）は除く
        If strFileName = ThisWorkbook.Name Then
            GoTo Continue
        End If
    
        ' エクセルファイルを開く
        fileAndPath = Path + "\" + strFileName
        Set Workbook = Workbooks.Open(fileAndPath)
    
        '一番先頭のシートから順にループ処理を行う
        For Each sheet In ActiveWorkbook.Sheets
            sheet.Activate                 '対象のシートをアクティブにする
            ActiveSheet.Range("A1").Select 'シートのA1を選択する
            ActiveWindow.Zoom = 100        '拡大倍率を100に設定する
        Next sheet
        Sheets(1).Select
    
        ' エクセルファイルを保存して閉じる
        Workbook.Save
        Workbook.Close
    
Continue:
    
        ' 次のファイル名を取得
        strFileName = Dir()
    Loop
End Sub

' /**
' * A1セルへのカーソル移動と拡大率100％のサブルーチンを
' * サブフォルダに対して再帰的に呼び出す。
' *
' * @Param Path 処理対象のフォルダパス
' *
' **/
Private Sub recursion(Path)
    Dim f As Object
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(Path).SubFolders
            Call setA1And100Per(f.Path)
        Next f
    End With
End Sub

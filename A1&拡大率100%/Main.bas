Attribute VB_Name = "Main"
Sub Main()
    Dim Path As String: Path = Constants.targetDir  '// 対象フォルダパス
    Dim Workbook As Workbook
    Dim sheet As Object
    
    Call beforeExecution
    
    ' フォルダの存在確認
    If Dir(Path, vbDirectory) = "" Then
        MsgBox "指定のフォルダは存在しません。", vbExclamation, Constants.dialogTitle
        Exit Sub
    End If
    '// 対象ファイルを編集する
    Call editFiles(Path)
    
    Call afterExecution
    MsgBox "完了しました。"

End Sub

'// 対象ファイルを編集する
Private Sub editFiles(Path)
    '// A1＆拡大率100％を設定する（ルートフォルダ）
    Call executeA1And100Per(Path)
    
    '// サブフォルダを再帰する
    If ThisWorkbook.Worksheets("Sheet1").CheckBoxes("CheckBox").Value = 1 Then
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
            If ThisWorkbook.Worksheets("Sheet1").CheckBoxes("CheckBox_font").Value = 1 Then
                ActiveSheet.Cells.font.Name = Constants.fontName()   'フォントを設定する
            End If
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
' * サブフォルダに対して再帰的に呼び出す。
' *
' * @Param Path 処理対象のフォルダパス
' *
' **/
Private Sub recursion(Path)
    Dim f As Object
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(Path).SubFolders
            Call editFiles(f.Path)
        Next f
    End With
End Sub

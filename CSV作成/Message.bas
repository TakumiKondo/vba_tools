Attribute VB_Name = "Message"
Function NotSelectedList() As String
    NotSelectedList = "1つ以上のファイルを選択してください。"
End Function

Function NotSelectedItem(target As String) As String
    NotSelectedItem = target & "を選択してください。"
End Function

Function completed() As String
    completed = "ファイル出力が完了しました。"
End Function

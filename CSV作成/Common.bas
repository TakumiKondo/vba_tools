Attribute VB_Name = "Common"
' 出力設定フォーム呼び出し
Sub showUserForm()
    UserForm1.Show
End Sub

' 出力対象シートの取得
Function getTargetSheets() As Object
    ' 除外するシート
    Dim exclusionSheets As Object
    Set exclusionSheets = CreateObject("Scripting.Dictionary")
    exclusionSheets.Add "要求・要件", 1
    exclusionSheets.Add "設計", 1
    exclusionSheets.Add "ファイル出力", 1
    
    Dim targetSheets As Object
    Set targetSheets = CreateObject("Scripting.Dictionary")
    
    ' 出力対象シート = 全シートから除外するシートを除いたもの
    For Each st In ThisWorkbook.Worksheets
        If Not exclusionSheets.Exists(st.Name) Then
            targetSheets.Add st.Name, 1
        End If
    Next st
    
    Set getTargetSheets = targetSheets
End Function

' 出力先パスの取得
Function outputPath() As String
    outputPath = ThisWorkbook.Path
End Function


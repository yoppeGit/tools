Attribute VB_Name = "module_ボタン押下"
Option Explicit

Sub 接続情報登録()
        
    '------------------------------
    ' 既存コネクション有無確認
    '------------------------------
    If PUB_sqlserverConnection <> "" Then
        If MsgBox("既存コネクションを上書きしますか？", vbOKCancel, "上書き確認") = vbCancel Then
            Exit Sub
        End If
    End If
    
    Set PUB_sqlserverConnection = Nothing
    form_接続情報登録.Show
    
    
End Sub


Sub SQL発行()
    '------------------------------
    ' 変数定義
    '------------------------------
    Dim selectedSQLFile As String
    Dim execSQLStatement() As String
    Dim sqlParts As Variant
    Dim addressDrawStart As String
    addressDrawStart = "B13"
    Dim addressDrawEnd As String
    
    '------------------------------
    ' コネクション確認
    '------------------------------
    If PUB_sqlserverConnection = "" Then
        MsgBox "接続情報を登録してください。"
        Call 接続情報登録
        Exit Sub
    End If
    
    '------------------------------
    ' SQLスクリプト選択
    '------------------------------
    ChDir ThisWorkbook.Path
    selectedSQLFile = Application.GetOpenFilename("SQLスクリプト,*.sql")
    
    '--- 選択チェック ---
    If selectedSQLFile = "False" Then
        MsgBox "ファイル選択がされませんでした。", vbOKOnly, "実行不可"
        Exit Sub
    End If
    
    
On Error GoTo TAG_ERROR

    '------------------------------
    ' SQLスクリプト文字列取得
    '------------------------------
    execSQLStatement = SQLスクリプト読み込み(selectedSQLFile)
    
    '------------------------------
    ' SQLスクリプト実行・結果描画
    '------------------------------
    For Each sqlParts In execSQLStatement
        If Len(sqlParts) = 0 Then
            GoTo Continue
        End If
        
        Call SQLServerスクリプト実行(sqlParts)
        
        addressDrawEnd = Select結果描画(addressDrawStart)
        
        addressDrawStart = Range(addressDrawEnd).Offset(2, 0).Address

Continue:
    Next sqlParts
    
    '------------------------------
    ' 終了処理
    '------------------------------
    MsgBox _
        "実行完了しました。" _
        , vbOKOnly _
        , "実行完了：SQL発行"
    
    Exit Sub
    
TAG_ERROR:
    MsgBox _
        "実行中にエラーが発生しました。内容を確認してください。" & vbCrLf & vbCrLf & _
        Err.Description _
        , vbCritical _
        , "エラー：SQL発行"
    
End Sub


Sub 接続情報破棄()
    Call SQLServer接続終了
End Sub



Sub 描画リセット()
    Rows("13:" & Cells.Rows.Count).Delete
End Sub


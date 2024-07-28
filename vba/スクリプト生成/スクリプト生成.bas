Attribute VB_Name = "スクリプト生成"
Option Explicit

'=============================================================
' スクリプト生成
' 選択パスに対しスクリプトを上書き保存する
'=============================================================
Function スクリプト上書き保存(ByVal filePath As String, ByVal writeStr As String, ByVal charCode As String)
    
    '----------------------------------------
    ' 変数
    '----------------------------------------
    
    
    
On Error GoTo TAG_ERROR
    
    '----------------------------------------
    ' ファイル書き込み
    '----------------------------------------
    With CreateObject("ADODB.Stream")
        .Charset = charCode
        .Open
        
        .WriteText writeStr
        
        .SaveToFile filePath, 2
        
        .Close
    End With
    
    MsgBox "以下ファイルへの書き込みが完了しました。" & vbCr & filePath
    
    Exit Function
    
    
TAG_ERROR:
    
    MsgBox "エラー発生しました。", vbCritical, "関数：スクリプト生成"
    
End Function

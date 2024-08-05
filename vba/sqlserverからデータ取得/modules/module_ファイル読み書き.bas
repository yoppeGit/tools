Attribute VB_Name = "module_ファイル読み書き"
Option Explicit

'=============================================================
' スクリプト読み込み
' 選択パスに対しスクリプトを読み込み文字列を返す
'=============================================================
Function SQLスクリプト読み込み(ByVal filePath As String) As String()
    
    '----------------------------------------
    ' 変数
    '----------------------------------------
    Dim statementDelimiter As String
    statementDelimiter = ";"
    Dim readText As String
    Dim returnStr() As String
    Dim objFile As Object
    
On Error GoTo TAG_ERROR
    
    '----------------------------------------
    ' ファイル書き込み
    '----------------------------------------
    With CreateObject("Scripting.FileSystemObject")
        Set objFile = .OpenTextFile(filePath)
        
        With objFile
            Do Until .AtEndOfStream
                readText = readText & .ReadLine
            Loop
        End With
        
    End With
    
    returnStr = Split(readText, statementDelimiter)
    
    SQLスクリプト読み込み = returnStr
    
    Exit Function
    
    
TAG_ERROR:
    
    MsgBox _
        "エラー発生しました。" & vbCrLf & Err.Description _
        , vbCritical _
        , "関数：スクリプト生成"
        
End Function


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


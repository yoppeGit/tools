Attribute VB_Name = "�X�N���v�g����"
Option Explicit

'=============================================================
' �X�N���v�g����
' �I���p�X�ɑ΂��X�N���v�g���㏑���ۑ�����
'=============================================================
Function �X�N���v�g�㏑���ۑ�(ByVal filePath As String, ByVal writeStr As String, ByVal charCode As String)
    
    '----------------------------------------
    ' �ϐ�
    '----------------------------------------
    
    
    
On Error GoTo TAG_ERROR
    
    '----------------------------------------
    ' �t�@�C����������
    '----------------------------------------
    With CreateObject("ADODB.Stream")
        .Charset = charCode
        .Open
        
        .WriteText writeStr
        
        .SaveToFile filePath, 2
        
        .Close
    End With
    
    MsgBox "�ȉ��t�@�C���ւ̏������݂��������܂����B" & vbCr & filePath
    
    Exit Function
    
    
TAG_ERROR:
    
    MsgBox "�G���[�������܂����B", vbCritical, "�֐��F�X�N���v�g����"
    
End Function

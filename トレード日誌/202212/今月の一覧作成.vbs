Option Explicit
Dim fso ,fol, sfol, ado, s

Set fso = WScript.CreateObject("Scripting.FileSystemObject")

' �t�H���_�ꗗ���擾
Set fol = fso.GetFolder(".")
s = "# �����̈ꗗ" & vbCrLf
For Each sfol In fol.SubFolders
    s = s & "[" & sfol.Name & "](./" & sfol.Name & "/main.md)" & vbCrLf
Next

' ���L���`�ǂݍ��݁����t������u��
Set ado = CreateObject("ADODB.Stream")
ado.Open
ado.Type = 2 ' �e�L�X�g�t�@�C��
ado.Charset = "UTF-8"
ado.LineSeparator = -1 ' ���s�R�[�h CRLF
ado.WriteText s, 1
ado.SaveToFile "main.md", 2
ado.Close

WScript.Echo "�쐬����"


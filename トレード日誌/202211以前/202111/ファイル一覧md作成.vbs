Option Explicit
Dim fso ,fol, f, ado, s

Set fso = WScript.CreateObject("Scripting.FileSystemObject")

' ファイル一覧を取得
Set fol = fso.GetFolder(".")
s = "# " & fol.Name & vbCrLf
For Each f In fol.Files
    If Right(LCase(f.Name), 4) = ".png" Then
        s = s & "## " & f.Name & vbCrLf
        s = s & "![](./" & f.Name & ")  " & vbCrLf
    End If
Next

Set ado = CreateObject("ADODB.Stream")
ado.Open
ado.Type = 2 ' テキストファイル
ado.Charset = "UTF-8"
ado.LineSeparator = -1 ' 改行コード CRLF
ado.WriteText s, 1
ado.SaveToFile "main.md", 2
ado.Close

WScript.Echo "作成完了"


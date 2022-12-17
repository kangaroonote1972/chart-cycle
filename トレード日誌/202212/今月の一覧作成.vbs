Option Explicit
Dim fso ,fol, sfol, ado, s

Set fso = WScript.CreateObject("Scripting.FileSystemObject")

' フォルダ一覧を取得
Set fol = fso.GetFolder(".")
s = "# 今月の一覧" & vbCrLf
For Each sfol In fol.SubFolders
    s = s & "[" & sfol.Name & "](./" & sfol.Name & "/main.md)" & vbCrLf
Next

' 日記雛形読み込み→日付部分を置換
Set ado = CreateObject("ADODB.Stream")
ado.Open
ado.Type = 2 ' テキストファイル
ado.Charset = "UTF-8"
ado.LineSeparator = -1 ' 改行コード CRLF
ado.WriteText s, 1
ado.SaveToFile "main.md", 2
ado.Close

WScript.Echo "作成完了"


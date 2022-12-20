createMonth = CDate("2023/01/01")
nextMonth = CDate("2023/02/01")
DSTR = "[YYYY/MM/DD 振り返り/次週想定](./YYYYMM/DD/main.md)  " & vbCrLf
str = ""
For I = 30 To 0 Step -1
	wMonth = DateAdd("d", I, createMonth)
	If wMonth <> nextMonth And Weekday(wMonth) <> vbSaturday Then
		y = Year(wMonth)
		m = Right("0" & CStr(Month(wMonth)), 2)
		d = Right("0" & CStr(Day(wMonth)), 2)
		s = Replace(DSTR, "YYYY", y)
		s = Replace(s, "MM", m)
		s = Replace(s, "DD", d)
		If Weekday(wMonth) <> vbSunday Then s = Replace(s, " 振り返り/次週想定", "")
		str = str & s
	End If
Next

Set ado = CreateObject("ADODB.Stream")
ado.Open
ado.Type = 2 ' テキストファイル
ado.Charset = "UTF-8"
ado.LineSeparator = -1 ' 改行コード CRLF
ado.WriteText str, 1
ado.SaveToFile "list.txt", 2
ado.Close

WScript.Echo "処理終了"

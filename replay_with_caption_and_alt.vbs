' 再生スクリプト (キャプションと代替テキストを同一入力)
Dim objFSO, objFile, WshShell
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")

' 記録ファイルを読み込む
If objFSO.FileExists("RecordedActions.txt") Then
    Set objFile = objFSO.OpenTextFile("RecordedActions.txt", 1)

    ' 記録されたURLを取得
    Do Until objFile.AtEndOfStream
        line = objFile.ReadLine

        If InStr(line, "URL: ") > 0 Then
            mainUrl = Replace(line, "URL: ", "")
            WshShell.Run "chrome.exe """ & mainUrl & """"
            WScript.Sleep 3000
        ElseIf InStr(line, "段落: ") > 0 Then
            paragraphText = Replace(line, "段落: ", "")
            WshShell.SendKeys paragraphText
            WshShell.SendKeys "{ENTER}"
            WScript.Sleep 500
        ElseIf InStr(line, "画像: ") > 0 Then
            imagePath = Replace(line, "画像: ", "")
            If imagePath <> "アップロードなし" Then
                If objFSO.FileExists(imagePath) Then
                    WshShell.SendKeys "%U"
                    WScript.Sleep 2000
                    WshShell.SendKeys imagePath
                    WshShell.SendKeys "{ENTER}"
                    WScript.Sleep 2000

                    ' キャプションを入力
                    captionText = Replace(objFile.ReadLine, "キャプション: ", "")
                    WshShell.SendKeys captionText
                    WshShell.SendKeys "{TAB}" ' フォーカスを代替テキストに移動
                    WshShell.SendKeys captionText ' 代替テキストとして同じ内容を入力
                    WshShell.SendKeys "{ENTER}"
                    WScript.Sleep 1000
                Else
                    MsgBox "指定された画像ファイルが見つかりません: " & imagePath, vbExclamation, "エラー"
                End If
            End If
        End If
    Loop

    objFile.Close
Else
    MsgBox "記録ファイルが見つかりません。"
End If

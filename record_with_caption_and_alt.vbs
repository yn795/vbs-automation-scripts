' 記録スクリプト (キャプションと代替テキストを同一入力)
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile("RecordedActions.txt", True)

Dim WshShell
Set WshShell = CreateObject("WScript.Shell")

' URLを入力
mainUrl = InputBox("記事編集画面のURLを入力してください:", "URLを入力")
If mainUrl = "" Then
    MsgBox "URLが入力されませんでした。スクリプトを終了します。"
    WScript.Quit
End If

' 記録開始
objFile.WriteLine "操作開始: " & Now
objFile.WriteLine "URL: " & mainUrl

' ブラウザを開く
WshShell.Run "chrome.exe """ & mainUrl & """"
WScript.Sleep 3000 ' ページの読み込み待ち

Do
    ' 段落のテキストを入力
    paragraphText = InputBox("段落のテキストを入力してください (終了するにはキャンセルを押してください):", "段落入力")
    If paragraphText = "" Then Exit Do

    objFile.WriteLine "段落: " & paragraphText
    WshShell.SendKeys paragraphText
    WshShell.SendKeys "{ENTER}" ' 次の段落へ移動
    WScript.Sleep 500

    ' 画像アップロード確認
    uploadImage = MsgBox("この段落に画像をアップロードしますか？", vbYesNo + vbQuestion, "画像アップロード確認")
    If uploadImage = vbYes Then
        Dim imagePath, captionText

        ' 画像パスを選択
        imagePath = InputBox("アップロードする画像のパスを入力してください:", "画像選択")
        If objFSO.FileExists(imagePath) Then
            objFile.WriteLine "画像: " & imagePath
            WshShell.SendKeys "%U" ' Alt + U でポップアップを開く (例)
            WScript.Sleep 2000
            WshShell.SendKeys imagePath
            WshShell.SendKeys "{ENTER}" ' 画像をアップロード
            WScript.Sleep 2000

            ' キャプションを入力 (デフォルト「あ」)
            captionText = InputBox("キャプションを入力してください:", "キャプション入力", "あ")
            objFile.WriteLine "キャプション: " & captionText
            WshShell.SendKeys captionText
            WshShell.SendKeys "{TAB}" ' フォーカスを代替テキストに移動
            WshShell.SendKeys captionText ' 代替テキストとして同じ内容を入力
            WshShell.SendKeys "{ENTER}" ' キャプション確定
            WScript.Sleep 1000
        Else
            MsgBox "指定された画像ファイルが見つかりません: " & imagePath, vbExclamation, "エラー"
        End If
    Else
        objFile.WriteLine "画像: アップロードなし"
    End If
Loop

objFile.WriteLine "操作終了: " & Now
objFile.Close

MsgBox "操作を記録しました: RecordedActions.txt"

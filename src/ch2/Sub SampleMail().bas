Sub SampleMail()
' メールアイテムオブジェクトの作成
    Dim objMailItem As Object
    Set objMailItem = CreateItem(olMailItem)
        With objMailItem
        '各種プロパティ設定
            .To = "testAddress@xx.xx" '正宛先
            .CC = "CCtestAddress@xx.xx" 'CC宛先
            .BCC = "BCCtestAddress@xx.xx" 'BCC宛先
            .Subject = "テスト件名" '件名
            '本文
            .Body = "テスト本文1⾏⽬" & vbCrLf & _
            "テスト本文2⾏⽬"
            'メールを表⽰する
            .Display '作成メールを表⽰
        End With
End Sub

Sub SampleTask()
    Dim objTaskItem As Object
    Set objTaskItem = CreateItem(olTaskItem)
        With objTaskItem
        .Subject = "テスト件名" '件名
        .StartDate = "2021/10/1" '開始日
        .DueDate = "2021/10/15" '期限
        .Importance = olImportanceNormal '優先度 2 olImportanceHigh 高 1 olImpotanceNormal 標準 0 olI,potanceLow 低
        .ReminderSet = True 'アラーム
            If .ReminderSet = True Then
                .ReminderTime = "2021/10/10 12:00" 'アラーム時間
            End If
        .Body = "テスト本文"
        .Display
    End With
End Sub

Sub SampleMailInfo()
'受信ボックスで選択している1番⽬のメールを抽出する
    Dim objSelect As Object
    Dim objMailItem As Object
    Set objSelect = ActiveExplorer.Selection
    Set objMailItem = objSelect.Item(1)
    'メールの情報をイミディエイトウィンドウに書き出す
        With objMailItem
            Debug.Print "受信⽇ :" & .ReceivedTime
            Debug.Print "送信者 :" & .SenderName
            Debug.Print "件名 :" & .Subject
            Debug.Print "送信先(To) :" & .To
            Debug.Print "送信先(CC) :" & .CC
            Debug.Print "送信先(BCC):" & .BCC
            Debug.Print "本文 :" & .Body
        End With
End Sub
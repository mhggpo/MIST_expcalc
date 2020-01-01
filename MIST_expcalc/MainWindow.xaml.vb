Class MainWindow
    Dim expfrom, expto, expall, needexp, onceexp, oncecost, onceexp1, onceexp2 As Long
    Dim times, elec As Integer
    Dim jianniangexp(91) As Long
    Dim tiduexp(151) As Long

    Private Sub button3_Click(sender As Object, e As RoutedEventArgs) Handles button3.Click
        MsgBox("接下来会打开本软件的Github页面，请按页面提示寻找软件最新版本哦！ 0v0", vbInformation, "昆昆提示")
        System.Diagnostics.Process.Start("https://github.com/mhggpo/MIST_expcalc")
    End Sub

    Private Sub type1_Checked(sender As Object, e As RoutedEventArgs) Handles type1.Checked
        textBoxto.Text = "90"
    End Sub

    Private Sub type2_Checked(sender As Object, e As RoutedEventArgs) Handles type2.Checked
        textBoxto.Text = "150"
    End Sub

    Private Sub button2_Click(sender As Object, e As RoutedEventArgs) Handles button2.Click
        End
    End Sub
    Private Sub button1_Click(sender As Object, e As RoutedEventArgs) Handles button1.Click
        MsgBox("程序开发:mhggpo 0v0" & vbCrLf & "经验表来源于NGA和QQ群内dalao,感谢dalao 0v0" & vbCrLf & "计算原理来自舰少资料库,感谢dalao 0v0" & vbCrLf & "软件最新版本参见Github 0v0", vbInformation, "昆昆提示")
    End Sub

    Dim showtext1, showtext2, showtext3, stage As String



    Public Sub setupexp(ByVal str1 As String, ByVal num1 As Long, ByVal num2 As Long)
        stage = str1
        onceexp = num1
        oncecost = num2
    End Sub


    Private Sub button_Click(sender As Object, e As RoutedEventArgs) Handles button.Click
        If type1.IsChecked = True Then
            jianniangexp = {0, 0, 10, 28, 55, 91, 135, 187, 248, 318, 396, 482, 611, 783, 997, 1253, 1552, 1894, 2278, 2704, 3173, 3685, 4281, 4963, 5729, 6581, 7517, 8539, 9645, 10837,
            12113, 13475, 14964, 16580, 18324, 20196, 22195, 24321, 26575, 28957, 31466, 34102, 36909, 39885, 43032, 46348, 49835, 53491, 57318, 61314, 65481, 69817,
            74579, 79765, 85377, 91413, 97875, 104761, 112073, 119809, 127971, 136557, 145994, 156280, 167417, 179403, 192240, 205926, 220463, 235849, 252086, 270022,
            289659, 310995, 334032, 358768, 385205, 413341, 443178, 474714, 507950, 542887, 580374, 620410, 662997, 708133, 755820, 806056, 858843, 914179, 972066}
            If Val(textBoxfrom.Text) >= 1 And Val(textBoxfrom.Text) <= Val(textBoxto.Text) Then
            Else
                MsgBox(“输入有误，请重新输入！”, vbInformation)
                textBoxfrom.Text = 1
                Exit Sub
            End If
            If Val(textBoxto.Text) <= 90 And Val(textBoxfrom.Text) <= Val(textBoxto.Text) Then
            Else
                MsgBox(“输入有误，请重新输入！”, vbInformation)
                textBoxto.Text = 90
                Exit Sub
            End If
            Dim level2 As Integer
            level2 = comboBox.SelectedIndex
            Select Case level2
                Case 0
                    setupexp("1-1N", 12, 8)
                Case 1
                    stage = "1-2N"
                    onceexp = 12
                    oncecost = 8
                Case 2
                    stage = "1-3N"
                    onceexp = 13
                    oncecost = 8
                Case 3
                    stage = "1-4N"
                    onceexp = 18
                    oncecost = 10
                Case 4
                    stage = "2-1N"
                    onceexp = 24
                    oncecost = 12
                Case 5
                    stage = "2-2N"
                    onceexp = 26
                    oncecost = 12
                Case 6
                    stage = "2-3N"
                    onceexp = 28
                    oncecost = 12
                Case 7
                    stage = "2-4N"
                    onceexp = 36
                    oncecost = 14
                Case 8
                    stage = "3-1N"
                    onceexp = 48
                    oncecost = 16
                Case 9
                    stage = "3-2N"
                    onceexp = 51
                    oncecost = 16
                Case 10
                    stage = "3-3N"
                    onceexp = 54
                    oncecost = 16
                Case 11
                    stage = "3-4N"
                    onceexp = 68
                    oncecost = 18
                Case 12
                    stage = "4-1N"
                    onceexp = 80
                    oncecost = 20
                Case 13
                    stage = "4-2N"
                    onceexp = 84
                    oncecost = 20
                Case 14
                    stage = "4-3N"
                    onceexp = 88
                    oncecost = 20
                Case 15
                    stage = "4-4N"
                    onceexp = 105
                    oncecost = 22
                Case 16
                    stage = "5-1N"
                    onceexp = 110
                    oncecost = 22
                Case 17
                    stage = "5-2N"
                    onceexp = 114
                    oncecost = 22
                Case 18
                    stage = "5-3N"
                    onceexp = 118
                    oncecost = 22
                Case 19
                    stage = "5-4N"
                    onceexp = 133
                    oncecost = 23
                Case 20
                    stage = "6-1N"
                    onceexp = 144
                    oncecost = 24
                Case 21
                    stage = "6-2N"
                    onceexp = 148
                    oncecost = 24
                Case 22
                    stage = "6-3N"
                    onceexp = 153
                    oncecost = 24
                Case 23
                    stage = "6-4N"
                    onceexp = 176
                    oncecost = 26
                Case 24
                    setupexp("7-1N", 182, 26)
                Case 25
                    setupexp("7-2N", 204, 28)
                Case 26
                    setupexp("7-3N", 212, 28)
                Case 27
                    setupexp("7-4N", 240, 30)
                Case 28
                    stage = "1-1H"
                    onceexp = 21
                    oncecost = 12
                Case 29
                    stage = "1-2H"
                    onceexp = 23
                    oncecost = 12
                Case 30
                    stage = "1-3H"
                    onceexp = 24
                    oncecost = 12
                Case 31
                    stage = "1-4H"
                    onceexp = 30
                    oncecost = 14
                Case 32
                    stage = "2-1H"
                    onceexp = 38
                    oncecost = 16
                Case 33
                    stage = "2-2H"
                    onceexp = 42
                    oncecost = 16
                Case 34
                    stage = "2-3H"
                    onceexp = 46
                    oncecost = 16
                Case 35
                    stage = "2-4H"
                    onceexp = 60
                    oncecost = 18
                Case 36
                    stage = "3-1H"
                    onceexp = 72
                    oncecost = 20
                Case 37
                    stage = "3-2H"
                    onceexp = 76
                    oncecost = 20
                Case 38
                    stage = "3-3H"
                    onceexp = 81
                    oncecost = 20
                Case 39
                    stage = "3-4H"
                    onceexp = 100
                    oncecost = 22
                Case 40
                    stage = "4-1H"
                    onceexp = 115
                    oncecost = 24
                Case 41
                    stage = "4-2H"
                    onceexp = 120
                    oncecost = 24
                Case 42
                    stage = "4-3H"
                    onceexp = 126
                    oncecost = 24
                Case 43
                    stage = "4-4H"
                    onceexp = 149
                    oncecost = 26
                Case 44
                    stage = "5-1H"
                    onceexp = 156
                    oncecost = 26
                Case 45
                    stage = "5-2H"
                    onceexp = 162
                    oncecost = 26
                Case 46
                    stage = "5-3H"
                    onceexp = 168
                    oncecost = 26
                Case 47
                    stage = "5-4H"
                    onceexp = 187
                    oncecost = 27
                Case 48
                    stage = "6-1H"
                    onceexp = 201
                    oncecost = 28
                Case 49
                    stage = "6-2H"
                    onceexp = 208
                    oncecost = 28
                Case 50
                    stage = "6-3H"
                    onceexp = 215
                    oncecost = 28
                Case 51
                    stage = "6-4H"
                    onceexp = 244
                    oncecost = 30
                Case 52
                    setupexp("7-1H", 247, 30)
                Case 53
                    setupexp("7-2H", 255, 30)
                Case 54
                    setupexp("7-3H", 263, 30)
                Case 55
                    setupexp("7-4H", 314, 34)
                Case 56
                    stage = "训练海域 Lv.10"
                    onceexp = 50
                    oncecost = 1
                Case 57
                    stage = "训练海域 Lv.20"
                    onceexp = 100
                    oncecost = 1
                Case 58
                    stage = "训练海域 Lv.40"
                    onceexp = 200
                    oncecost = 1
                Case 59
                    stage = "训练海域 Lv.60"
                    onceexp = 400
                    oncecost = 1
                Case 60
                    stage = "自定义"
                    oncecost = 1
                    Dim temp1 As String
Input1:
                    temp1 = InputBox("请输入关卡经验", "昆昆提示")
                    If temp1 = "" Then MsgBox("哼，居然想用空白的条件来忽悠昆西╭(╯^╰)╮，不可原谅！", vbInformation) : GoTo Input1
                    If temp1 = "0" Then MsgBox("昆昆才不会被骗去吃土呢！", vbInformation) : GoTo Input1
                    onceexp = Val(temp1)
                    If onceexp = 0 Then MsgBox("昆昆才不会被骗去吃土呢！", vbInformation) : GoTo Input1
            End Select
            expfrom = jianniangexp(CInt(textBoxfrom.Text))
            expto = jianniangexp(CInt(textBoxto.Text))
            needexp = expto - expfrom
            If (addexp.IsChecked = True) Then
                expall = 1600 * Val(expa.Text) + 400 * Val(expb.Text) + 100 * Val(expc.Text)
            Else
                expall = 0
            End If
            Dim tempforexp As Long
            tempforexp = 0
            If (addexp.IsChecked = True) Then
                tempforexp = expto - expfrom - expall
            End If
            If (tempforexp > 0 And addexp.IsChecked = True) Then
                times = (tempforexp / onceexp)
            ElseIf (tempforexp <= 0 And addexp.IsChecked = True) Then
                While (tempforexp <= 0)
                    MsgBox("昆西稍稍思考，发现当前条件下经验溢出，将自动提升一个目标等级计算！", vbInformation, "昆昆提示")
                    textBoxto.Text = textBoxto.Text + 1
                    If (textBoxto.Text > 90) Then
                        MsgBox("目标等级大于上限！", vbCritical, "昆昆提示")
                        Exit Sub
                    End If
                    expto = jianniangexp(CInt(textBoxto.Text))
                    needexp = expto - expfrom
                    tempforexp = expto - expfrom - expall
                End While
                times = (tempforexp / onceexp)
            Else times = (needexp / onceexp)
            End If
            elec = times * oncecost
            If level2 >= 56 And level2 <= 59 Then
                MsgBox("充满智慧的昆西掐指一算：" & "舰娘由等级" & textBoxfrom.Text & "升级至" & textBoxto.Text & "，共需经验值" & needexp & vbCrLf & "刷" & stage & "关卡的单个节点" & vbCrLf & "一次经验" & onceexp & "，需要次数：" & times & " 需要深海抑制装置：" & elec, vbInformation, "昆昆提示")
            ElseIf level2 = 60 Then
                MsgBox("充满智慧的昆西掐指一算：" & "舰娘由等级" & textBoxfrom.Text & "升级至" & textBoxto.Text & "，共需经验值" & needexp & vbCrLf & "刷" & stage & "关卡的单个节点" & vbCrLf & "一次经验" & onceexp & "，需要次数：" & times, vbInformation, "昆昆提示")
            Else
                MsgBox("充满智慧的昆西掐指一算：" & "舰娘由等级" & textBoxfrom.Text & "升级至" & textBoxto.Text & "，共需经验值" & needexp & vbCrLf & "刷" & stage & "关卡的单个节点" & vbCrLf & "一次经验" & onceexp & "，需要次数：" & times & " 需要电：" & elec, vbInformation, "昆昆提示")
                End If
            Else
            tiduexp = {0, 0, 10, 30, 60, 100, 150, 210, 280, 360, 450, 550, 670, 810, 970, 1150, 1350, 1570, 1810, 2070, 2350, 2650, 2980, 3340, 3730, 4150, 4600, 5080, 5590, 6130, 6700,
            7300, 7940, 8620, 9340, 10100, 10900, 11740, 12620, 13540, 14500, 15500, 16550, 17650, 18800, 20000, 21250, 22550, 23900, 25300, 26750, 28250, 29810, 31430,
            33110, 34850, 36650, 38510, 40430, 42410, 44450, 46550, 48720, 50960, 53270, 55650, 58100, 60620, 63210, 65870, 68600, 71400, 74280, 77240, 80280, 83400,
            86600, 89880, 93240, 96680, 100200, 103800, 107490, 111270, 115140, 119100, 123150, 127290, 131520, 135840, 140250, 144750, 149350, 154050, 158850, 163750,
            168750, 173850, 179050, 184350, 189750, 195250, 200850, 206550, 212350, 218250, 224250, 230350, 236550, 242850, 249250, 255750, 262350, 269050, 275850,
            282750, 289750, 296850, 304050, 311350, 318750, 326250, 333850, 341550, 349350, 357250, 365250, 373350, 381550, 389850, 398250, 406750, 415350, 424050,
            432850, 441750, 450750, 459850, 469050, 478350, 487750, 497250, 506850, 516550, 526350, 536250, 546250, 556350, 566550, 576850, 587250}
            If Val(textBoxfrom.Text) >= 1 And Val(textBoxfrom.Text) <= Val(textBoxto.Text) Then
            Else
                MsgBox(“输入有误，请重新输入！”, vbInformation)
                textBoxfrom.Text = 1
                Exit Sub
            End If
            If Val(textBoxto.Text) <= 150 And Val(textBoxfrom.Text) <= Val(textBoxto.Text) Then
            Else
                MsgBox(“输入有误，请重新输入！”, vbInformation)
                textBoxto.Text = 150
                Exit Sub
            End If
            Dim level2 As Integer
            level2 = comboBox.SelectedIndex
            Select Case level2
                Case 0
                    setupexp("1-1N", 12, 15)
                Case 1
                    stage = "1-2N"
                    onceexp = 12
                    oncecost = 8
                Case 2
                    stage = "1-3N"
                    onceexp = 13
                    oncecost = 8
                Case 3
                    stage = "1-4N"
                    onceexp = 18
                    oncecost = 10
                Case 4
                    stage = "2-1N"
                    onceexp = 24
                    oncecost = 12
                Case 5
                    stage = "2-2N"
                    onceexp = 26
                    oncecost = 12
                Case 6
                    stage = "2-3N"
                    onceexp = 28
                    oncecost = 12
                Case 7
                    stage = "2-4N"
                    onceexp = 36
                    oncecost = 14
                Case 8
                    stage = "3-1N"
                    onceexp = 48
                    oncecost = 16
                Case 9
                    stage = "3-2N"
                    onceexp = 51
                    oncecost = 16
                Case 10
                    stage = "3-3N"
                    onceexp = 54
                    oncecost = 16
                Case 11
                    stage = "3-4N"
                    onceexp = 68
                    oncecost = 18
                Case 12
                    stage = "4-1N"
                    onceexp = 80
                    oncecost = 20
                Case 13
                    stage = "4-2N"
                    onceexp = 84
                    oncecost = 20
                Case 14
                    stage = "4-3N"
                    onceexp = 88
                    oncecost = 20
                Case 15
                    stage = "4-4N"
                    onceexp = 105
                    oncecost = 22
                Case 16
                    stage = "5-1N"
                    onceexp = 110
                    oncecost = 22
                Case 17
                    stage = "5-2N"
                    onceexp = 114
                    oncecost = 22
                Case 18
                    stage = "5-3N"
                    onceexp = 118
                    oncecost = 22
                Case 19
                    stage = "5-4N"
                    onceexp = 133
                    oncecost = 23
                Case 20
                    stage = "6-1N"
                    onceexp = 144
                    oncecost = 24
                Case 21
                    stage = "6-2N"
                    onceexp = 148
                    oncecost = 24
                Case 22
                    stage = "6-3N"
                    onceexp = 153
                    oncecost = 24
                Case 23
                    stage = "6-4N"
                    onceexp = 176
                    oncecost = 26
                Case 24
                    setupexp("7-1N", 182, 26)
                Case 25
                    setupexp("7-2N", 204, 28)
                Case 26
                    setupexp("7-3N", 212, 28)
                Case 27
                    setupexp("7-4N", 240, 30)
                Case 28
                    stage = "1-1H"
                    onceexp = 21
                    oncecost = 12
                Case 29
                    stage = "1-2H"
                    onceexp = 23
                    oncecost = 12
                Case 30
                    stage = "1-3H"
                    onceexp = 24
                    oncecost = 12
                Case 31
                    stage = "1-4H"
                    onceexp = 30
                    oncecost = 14
                Case 32
                    stage = "2-1H"
                    onceexp = 38
                    oncecost = 16
                Case 33
                    stage = "2-2H"
                    onceexp = 42
                    oncecost = 16
                Case 34
                    stage = "2-3H"
                    onceexp = 46
                    oncecost = 16
                Case 35
                    stage = "2-4H"
                    onceexp = 60
                    oncecost = 18
                Case 36
                    stage = "3-1H"
                    onceexp = 72
                    oncecost = 20
                Case 37
                    stage = "3-2H"
                    onceexp = 76
                    oncecost = 20
                Case 38
                    stage = "3-3H"
                    onceexp = 81
                    oncecost = 20
                Case 39
                    stage = "3-4H"
                    onceexp = 100
                    oncecost = 22
                Case 40
                    stage = "4-1H"
                    onceexp = 115
                    oncecost = 24
                Case 41
                    stage = "4-2H"
                    onceexp = 120
                    oncecost = 24
                Case 42
                    stage = "4-3H"
                    onceexp = 126
                    oncecost = 24
                Case 43
                    stage = "4-4H"
                    onceexp = 149
                    oncecost = 26
                Case 44
                    stage = "5-1H"
                    onceexp = 156
                    oncecost = 26
                Case 45
                    stage = "5-2H"
                    onceexp = 162
                    oncecost = 26
                Case 46
                    stage = "5-3H"
                    onceexp = 168
                    oncecost = 26
                Case 47
                    stage = "5-4H"
                    onceexp = 187
                    oncecost = 27
                Case 48
                    stage = "6-1H"
                    onceexp = 201
                    oncecost = 28
                Case 49
                    stage = "6-2H"
                    onceexp = 208
                    oncecost = 28
                Case 50
                    stage = "6-3H"
                    onceexp = 215
                    oncecost = 28
                Case 51
                    stage = "6-4H"
                    onceexp = 244
                    oncecost = 30
                Case 52
                    setupexp("7-1H", 247, 30)
                Case 53
                    setupexp("7-2H", 255, 30)
                Case 54
                    setupexp("7-3H", 263, 30)
                Case 55
                    setupexp("7-4H", 314, 34)
                Case 56
                    stage = "训练海域 Lv.10"
                    onceexp = 50
                    oncecost = 1
                Case 57
                    stage = "训练海域 Lv.20"
                    onceexp = 100
                    oncecost = 1
                Case 58
                    stage = "训练海域 Lv.40"
                    onceexp = 200
                    oncecost = 1
                Case 59
                    stage = "训练海域 Lv.60"
                    onceexp = 400
                    oncecost = 1
                Case 60
                    stage = "自定义"
                    oncecost = 1
                    Dim temp1 As String
Input3:
                    temp1 = InputBox("请输入关卡经验", "昆昆提示")
                    If temp1 = "" Then MsgBox("哼，居然想用空白的条件来忽悠昆西╭(╯^╰)╮，不可原谅！", vbInformation) : GoTo Input3
                    If temp1 = "0" Then MsgBox("昆昆才不会被骗去吃土呢！", vbInformation) : GoTo Input3
                    onceexp = Val(temp1)
                    If onceexp = 0 Then MsgBox("昆昆才不会被骗去吃土呢！", vbInformation) : GoTo Input3
            End Select
            If level2 >= 0 And level2 <= 55 Then onceexp = oncecost
            expfrom = tiduexp(CInt(textBoxfrom.Text))
            expto = tiduexp(CInt(textBoxto.Text))
            needexp = expto - expfrom
            times = needexp / onceexp
            If times = 0 Then times = times + 1
            elec = times * oncecost
            If level2 >= 56 And level2 <= 59 Then
                MsgBox("充满智慧的昆西掐指一算：" & "提督由等级" & textBoxfrom.Text & "升级至" & textBoxto.Text & "，共需经验值" & needexp & vbCrLf & "刷" & stage & "关卡的单个节点" & vbCrLf & "一次经验" & onceexp & "，需要次数：" & times & " 需要深海抑制装置：" & elec, vbInformation, "昆昆提示")
            ElseIf level2 = 60 Then
                MsgBox("充满智慧的昆西掐指一算：" & "提督由等级" & textBoxfrom.Text & "升级至" & textBoxto.Text & "，共需经验值" & needexp & vbCrLf & "刷" & stage & "关卡的单个节点" & vbCrLf & "一次经验" & onceexp & "，需要次数：" & times, vbInformation, "昆昆提示")
            Else
                MsgBox("充满智慧的昆西掐指一算：" & "提督由等级" & textBoxfrom.Text & "升级至" & textBoxto.Text & "，共需经验值" & needexp & vbCrLf & "刷" & stage & "关卡的单个节点" & vbCrLf & "一次经验" & onceexp & "，需要次数：" & times & " 需要电：" & elec, vbInformation, "昆昆提示")
            End If
        End If
    End Sub
End Class
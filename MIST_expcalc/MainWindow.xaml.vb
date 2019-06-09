Class MainWindow
    Dim expfrom, expto, needexp, onceexp, oncecost, onceexp1, onceexp2 As Long
    Dim times, elec As Integer
    Dim jianniangexp(81) As Long
    Dim tiduexp(151) As Long

    Private Sub button3_Click(sender As Object, e As RoutedEventArgs) Handles button3.Click
        MsgBox("接下来会打开本软件的Github页面，请按页面提示寻找软件最新版本哦！ 0v0", vbInformation, "昆昆提示")
        System.Diagnostics.Process.Start("https://github.com/mhggpo/MIST_expcalc")
    End Sub

    Private Sub type1_Checked(sender As Object, e As RoutedEventArgs) Handles type1.Checked
        textBoxto.Text = "80"
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







    Private Sub button_Click(sender As Object, e As RoutedEventArgs) Handles button.Click
        If type1.IsChecked = True Then
            jianniangexp(0) = 0
            jianniangexp(1) = 0
            jianniangexp(2) = 10
            jianniangexp(3) = 28
            jianniangexp(4) = 55
            jianniangexp(5) = 91
            jianniangexp(6) = 135
            jianniangexp(7) = 187
            jianniangexp(8) = 248
            jianniangexp(9) = 318
            jianniangexp(10) = 396
            jianniangexp(11) = 482
            jianniangexp(12) = 611
            jianniangexp(13) = 783
            jianniangexp(14) = 997
            jianniangexp(15) = 1253
            jianniangexp(16) = 1552
            jianniangexp(17) = 1894
            jianniangexp(18) = 2278
            jianniangexp(19) = 2704
            jianniangexp(20) = 3173
            jianniangexp(21) = 3685
            jianniangexp(22) = 4281
            jianniangexp(23) = 4963
            jianniangexp(24) = 5729
            jianniangexp(25) = 6581
            jianniangexp(26) = 7517
            jianniangexp(27) = 8539
            jianniangexp(28) = 9645
            jianniangexp(29) = 10837
            jianniangexp(30) = 12113
            jianniangexp(31) = 13475
            jianniangexp(32) = 14964
            jianniangexp(33) = 16580
            jianniangexp(34) = 18324
            jianniangexp(35) = 20196
            jianniangexp(36) = 22195
            jianniangexp(37) = 24321
            jianniangexp(38) = 26575
            jianniangexp(39) = 28957
            jianniangexp(40) = 31466
            jianniangexp(41) = 34102
            jianniangexp(42) = 36909
            jianniangexp(43) = 39885
            jianniangexp(44) = 43032
            jianniangexp(45) = 46348
            jianniangexp(46) = 49835
            jianniangexp(47) = 53491
            jianniangexp(48) = 57318
            jianniangexp(49) = 61314
            jianniangexp(50) = 65481
            jianniangexp(51) = 69817
            jianniangexp(52) = 74579
            jianniangexp(53) = 79765
            jianniangexp(54) = 85377
            jianniangexp(55) = 91413
            jianniangexp(56) = 97875
            jianniangexp(57) = 104761
            jianniangexp(58) = 112073
            jianniangexp(59) = 119809
            jianniangexp(60) = 127971
            jianniangexp(61) = 136557
            jianniangexp(62) = 145994
            jianniangexp(63) = 156280
            jianniangexp(64) = 167417
            jianniangexp(65) = 179403
            jianniangexp(66) = 192240
            jianniangexp(67) = 205926
            jianniangexp(68) = 220463
            jianniangexp(69) = 235849
            jianniangexp(70) = 252086
            jianniangexp(71) = 270022
            jianniangexp(72) = 289659
            jianniangexp(73) = 310995
            jianniangexp(74) = 334032
            jianniangexp(75) = 358768
            jianniangexp(76) = 385205
            jianniangexp(77) = 413341
            jianniangexp(78) = 443178
            jianniangexp(79) = 474714
            jianniangexp(80) = 507951
            If Val(textBoxfrom.Text) >= 1 And Val(textBoxfrom.Text) <= Val(textBoxto.Text) Then
            Else
                MsgBox(“输入有误，请重新输入！”, vbInformation)
                textBoxfrom.Text = 1
                Exit Sub
            End If
            If Val(textBoxto.Text) <= 80 And Val(textBoxfrom.Text) <= Val(textBoxto.Text) Then
            Else
                MsgBox(“输入有误，请重新输入！”, vbInformation)
                textBoxto.Text = 80
                Exit Sub
            End If
            Dim level2 As Integer
            level2 = comboBox.SelectedIndex
            Select Case level2
                Case 0
                    stage = "1-1N"
                    onceexp = 12
                    oncecost = 8
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
                    stage = "1-1H"
                    onceexp = 21
                    oncecost = 12
                Case 21
                    stage = "1-2H"
                    onceexp = 23
                    oncecost = 12
                Case 22
                    stage = "1-3H"
                    onceexp = 24
                    oncecost = 12
                Case 23
                    stage = "1-4H"
                    onceexp = 30
                    oncecost = 14
                Case 24
                    stage = "2-1H"
                    onceexp = 38
                    oncecost = 16
                Case 25
                    stage = "2-2H"
                    onceexp = 42
                    oncecost = 16
                Case 26
                    stage = "2-3H"
                    onceexp = 46
                    oncecost = 16
                Case 27
                    stage = "2-4H"
                    onceexp = 60
                    oncecost = 18
                Case 28
                    stage = "3-1H"
                    onceexp = 72
                    oncecost = 20
                Case 29
                    stage = "3-2H"
                    onceexp = 76
                    oncecost = 20
                Case 30
                    stage = "3-3H"
                    onceexp = 81
                    oncecost = 20
                Case 31
                    stage = "3-4H"
                    onceexp = 100
                    oncecost = 22
                Case 32
                    stage = "4-1H"
                    onceexp = 115
                    oncecost = 24
                Case 33
                    stage = "4-2H"
                    onceexp = 120
                    oncecost = 24
                Case 34
                    stage = "4-3H"
                    onceexp = 126
                    oncecost = 24
                Case 35
                    stage = "4-4H"
                    onceexp = 149
                    oncecost = 26
                Case 36
                    stage = "5-1H"
                    onceexp = 156
                    oncecost = 26
                Case 37
                    stage = "5-2H"
                    onceexp = 162
                    oncecost = 26
                Case 38
                    stage = "5-3H"
                    onceexp = 168
                    oncecost = 26
                Case 39
                    stage = "5-4H"
                    onceexp = 187
                    oncecost = 27
                Case 40
                    stage = "训练海域 Lv.10"
                    onceexp = 50
                    oncecost = 1
                Case 41
                    stage = "训练海域 Lv.20"
                    onceexp = 100
                    oncecost = 1
                Case 42
                    stage = "训练海域 Lv.40"
                    onceexp = 200
                    oncecost = 1
                Case 43
                    stage = "训练海域 Lv.60"
                    onceexp = 400
                    oncecost = 1
                Case 44
                    stage = "自定义"
                    oncecost = 1
                    Dim temp1 As String
Input1:
                    temp1 = InputBox("请输入关卡经验", "昆昆提示")
                    If temp1 = "" Then MsgBox("哼，居然想用空白的条件来忽悠昆西╭(╯^╰)╮，不可原谅！", vbInformation) : GoTo Input1
                    If temp1 = "0" Then MsgBox("昆昆才不会被骗去1-5吃土呢！", vbInformation) : GoTo Input1
                    onceexp = Val(temp1)
                    If onceexp = 0 Then MsgBox("昆昆才不会被骗去1-5吃土呢！", vbInformation) : GoTo Input1
                Case 45
                    stage = "6-1N"
                    onceexp = 144
                    oncecost = 24
                Case 46
                    stage = "6-2N"
                    onceexp = 148
                    oncecost = 24
                Case 47
                    stage = "6-3N"
                    onceexp = 153
                    oncecost = 24
                Case 48
                    stage = "6-4N"
                    onceexp = 176
                    oncecost = 26
                Case 49
                    stage = "6-1H"
                    onceexp = 201
                    oncecost = 28
                Case 50
                    stage = "6-2H"
                    onceexp = 208
                    oncecost = 28
                Case 51
                    stage = "6-3H"
                    onceexp = 215
                    oncecost = 28
                Case 52
                    stage = "6-4H"
                    onceexp = 244
                    oncecost = 30
            End Select
            expfrom = jianniangexp(CInt(textBoxfrom.Text))
            expto = jianniangexp(CInt(textBoxto.Text))
            needexp = expto - expfrom
            Dim tempforke As Long
            tempforke = onceexp
            times = needexp / onceexp
            elec = times * oncecost
            If level2 >= 40 And level2 <= 43 Then
                MsgBox("充满智慧的昆西掐指一算：" & "舰娘由等级" & textBoxfrom.Text & "升级至" & textBoxto.Text & "，共需经验值" & needexp & vbCrLf & "刷" & stage & "关卡的单个节点" & vbCrLf & "一次经验" & onceexp & "，需要次数：" & times & " 需要深海抑制装置：" & elec, vbInformation, "昆昆提示")
            ElseIf level2 = 44 Then
                MsgBox("充满智慧的昆西掐指一算：" & "舰娘由等级" & textBoxfrom.Text & "升级至" & textBoxto.Text & "，共需经验值" & needexp & vbCrLf & "刷" & stage & "关卡的单个节点" & vbCrLf & "一次经验" & onceexp & "，需要次数：" & times, vbInformation, "昆昆提示")
            Else
                MsgBox("充满智慧的昆西掐指一算：" & "舰娘由等级" & textBoxfrom.Text & "升级至" & textBoxto.Text & "，共需经验值" & needexp & vbCrLf & "刷" & stage & "关卡的单个节点" & vbCrLf & "一次经验" & onceexp & "，需要次数：" & times & " 需要电：" & elec, vbInformation, "昆昆提示")
            End If
        Else
            tiduexp(0) = 0
            tiduexp(1) = 0
            tiduexp(2) = 10
            tiduexp(3) = 30
            tiduexp(4) = 60
            tiduexp(5) = 100
            tiduexp(6) = 150
            tiduexp(7) = 210
            tiduexp(8) = 280
            tiduexp(9) = 360
            tiduexp(10) = 450
            tiduexp(11) = 550
            tiduexp(12) = 670
            tiduexp(13) = 810
            tiduexp(14) = 970
            tiduexp(15) = 1150
            tiduexp(16) = 1350
            tiduexp(17) = 1570
            tiduexp(18) = 1810
            tiduexp(19) = 2070
            tiduexp(20) = 2350
            tiduexp(21) = 2650
            tiduexp(22) = 2980
            tiduexp(23) = 3340
            tiduexp(24) = 3730
            tiduexp(25) = 4150
            tiduexp(26) = 4600
            tiduexp(27) = 5080
            tiduexp(28) = 5590
            tiduexp(29) = 6130
            tiduexp(30) = 6700
            tiduexp(31) = 7300
            tiduexp(32) = 7940
            tiduexp(33) = 8620
            tiduexp(34) = 9340
            tiduexp(35) = 10100
            tiduexp(36) = 10900
            tiduexp(37) = 11740
            tiduexp(38) = 12620
            tiduexp(39) = 13540
            tiduexp(40) = 14500
            tiduexp(41) = 15500
            tiduexp(42) = 16550
            tiduexp(43) = 17650
            tiduexp(44) = 18800
            tiduexp(45) = 20000
            tiduexp(46) = 21250
            tiduexp(47) = 22550
            tiduexp(48) = 23900
            tiduexp(49) = 25300
            tiduexp(50) = 26750
            tiduexp(51) = 28250
            tiduexp(52) = 29810
            tiduexp(53) = 31430
            tiduexp(54) = 33110
            tiduexp(55) = 34850
            tiduexp(56) = 36650
            tiduexp(57) = 38510
            tiduexp(58) = 40430
            tiduexp(59) = 42410
            tiduexp(60) = 44450
            tiduexp(61) = 46550
            tiduexp(62) = 48720
            tiduexp(63) = 50960
            tiduexp(64) = 53270
            tiduexp(65) = 55650
            tiduexp(66) = 58100
            tiduexp(67) = 60620
            tiduexp(68) = 63210
            tiduexp(69) = 65870
            tiduexp(70) = 68600
            tiduexp(71) = 71400
            tiduexp(72) = 74280
            tiduexp(73) = 77240
            tiduexp(74) = 80280
            tiduexp(75) = 83400
            tiduexp(76) = 86600
            tiduexp(77) = 89880
            tiduexp(78) = 93240
            tiduexp(79) = 96680
            tiduexp(80) = 100200
            tiduexp(81) = 103800
            tiduexp(82) = 107490
            tiduexp(83) = 111270
            tiduexp(84) = 115140
            tiduexp(85) = 119100
            tiduexp(86) = 123150
            tiduexp(87) = 127290
            tiduexp(88) = 131520
            tiduexp(89) = 135840
            tiduexp(90) = 140250
            tiduexp(91) = 144750
            tiduexp(92) = 149350
            tiduexp(93) = 154050
            tiduexp(94) = 158850
            tiduexp(95) = 163750
            tiduexp(96) = 168750
            tiduexp(97) = 173850
            tiduexp(98) = 179050
            tiduexp(99) = 184350
            tiduexp(100) = 189750
            tiduexp(101) = 195250
            tiduexp(102) = 200850
            tiduexp(103) = 206550
            tiduexp(104) = 212350
            tiduexp(105) = 218250
            tiduexp(106) = 224250
            tiduexp(107) = 230350
            tiduexp(108) = 236550
            tiduexp(109) = 242850
            tiduexp(110) = 249250
            tiduexp(111) = 255750
            tiduexp(112) = 262350
            tiduexp(113) = 269050
            tiduexp(114) = 275850
            tiduexp(115) = 282750
            tiduexp(116) = 289750
            tiduexp(117) = 296850
            tiduexp(118) = 304050
            tiduexp(119) = 311350
            tiduexp(120) = 318750
            tiduexp(121) = 326250
            tiduexp(122) = 333850
            tiduexp(123) = 341550
            tiduexp(124) = 349350
            tiduexp(125) = 357250
            tiduexp(126) = 365250
            tiduexp(127) = 373350
            tiduexp(128) = 381550
            tiduexp(129) = 389850
            tiduexp(130) = 398250
            tiduexp(131) = 406750
            tiduexp(132) = 415350
            tiduexp(133) = 424050
            tiduexp(134) = 432850
            tiduexp(135) = 441750
            tiduexp(136) = 450750
            tiduexp(137) = 459850
            tiduexp(138) = 469050
            tiduexp(139) = 478350
            tiduexp(140) = 487750
            tiduexp(141) = 497250
            tiduexp(142) = 506850
            tiduexp(143) = 516550
            tiduexp(144) = 526350
            tiduexp(145) = 536250
            tiduexp(146) = 546250
            tiduexp(147) = 556350
            tiduexp(148) = 566550
            tiduexp(149) = 576850
            tiduexp(150) = 587250
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
                    stage = "1-1N"
                    onceexp = 8
                    oncecost = 8
                Case 1
                    stage = "1-2N"
                    onceexp = 8
                    oncecost = 8
                Case 2
                    stage = "1-3N"
                    onceexp = 8
                    oncecost = 8
                Case 3
                    stage = "1-4N"
                    onceexp = 10
                    oncecost = 10
                Case 4
                    stage = "2-1N"
                    onceexp = 12
                    oncecost = 12
                Case 5
                    stage = "2-2N"
                    onceexp = 12
                    oncecost = 12
                Case 6
                    stage = "2-3N"
                    onceexp = 12
                    oncecost = 12
                Case 7
                    stage = "2-4N"
                    onceexp = 14
                    oncecost = 14
                Case 8
                    stage = "3-1N"
                    onceexp = 16
                    oncecost = 16
                Case 9
                    stage = "3-2N"
                    onceexp = 16
                    oncecost = 16
                Case 10
                    stage = "3-3N"
                    onceexp = 16
                    oncecost = 16
                Case 11
                    stage = "3-4N"
                    onceexp = 18
                    oncecost = 18
                Case 12
                    stage = "4-1N"
                    onceexp = 20
                    oncecost = 20
                Case 13
                    stage = "4-2N"
                    onceexp = 20
                    oncecost = 20
                Case 14
                    stage = "4-3N"
                    onceexp = 20
                    oncecost = 20
                Case 15
                    stage = "4-4N"
                    onceexp = 22
                    oncecost = 22
                Case 16
                    stage = "5-1N"
                    onceexp = 22
                    oncecost = 22
                Case 17
                    stage = "5-2N"
                    onceexp = 22
                    oncecost = 22
                Case 18
                    stage = "5-3N"
                    onceexp = 22
                    oncecost = 22
                Case 19
                    stage = "5-4N"
                    onceexp = 23
                    oncecost = 23
                Case 20
                    stage = "1-1H"
                    onceexp = 12
                    oncecost = 12
                Case 21
                    stage = "1-2H"
                    onceexp = 12
                    oncecost = 12
                Case 22
                    stage = "1-3H"
                    onceexp = 12
                    oncecost = 12
                Case 23
                    stage = "1-4H"
                    onceexp = 14
                    oncecost = 14
                Case 24
                    stage = "2-1H"
                    onceexp = 16
                    oncecost = 16
                Case 25
                    stage = "2-2H"
                    onceexp = 16
                    oncecost = 16
                Case 26
                    stage = "2-3H"
                    onceexp = 16
                    oncecost = 16
                Case 27
                    stage = "2-4H"
                    onceexp = 18
                    oncecost = 18
                Case 28
                    stage = "3-1H"
                    onceexp = 20
                    oncecost = 20
                Case 29
                    stage = "3-2H"
                    onceexp = 20
                    oncecost = 20
                Case 30
                    stage = "3-3H"
                    onceexp = 20
                    oncecost = 20
                Case 31
                    stage = "3-4H"
                    onceexp = 22
                    oncecost = 22
                Case 32
                    stage = "4-1H"
                    onceexp = 24
                    oncecost = 24
                Case 33
                    stage = "4-2H"
                    onceexp = 24
                    oncecost = 24
                Case 34
                    stage = "4-3H"
                    onceexp = 24
                    oncecost = 24
                Case 35
                    stage = "4-4H"
                    onceexp = 26
                    oncecost = 26
                Case 36
                    stage = "5-1H"
                    onceexp = 26
                    oncecost = 26
                Case 37
                    stage = "5-2H"
                    onceexp = 162
                    oncecost = 26
                Case 38
                    stage = "5-3H"
                    onceexp = 26
                    oncecost = 26
                Case 39
                    stage = "5-4H"
                    onceexp = 27
                    oncecost = 27
                Case 40
                    stage = "训练海域 Lv.10"
                    onceexp = 10
                    oncecost = 1
                Case 41
                    stage = "训练海域 Lv.20"
                    onceexp = 20
                    oncecost = 1
                Case 42
                    stage = "训练海域 Lv.40"
                    onceexp = 40
                    oncecost = 1
                Case 43
                    stage = "训练海域 Lv.60"
                    onceexp = 80
                    oncecost = 1
                Case 44
                    stage = "自定义"
                    oncecost = 1
                    Dim temp1 As String
Input3:
                    temp1 = InputBox("请输入关卡经验", "昆昆提示")
                    If temp1 = "" Then MsgBox("哼，居然想用空白的条件来忽悠昆西╭(╯^╰)╮，不可原谅！", vbInformation) : GoTo Input3
                    If temp1 = "0" Then MsgBox("这样的条件下是没有提督经验的哦(⊙o⊙)", vbInformation) : GoTo Input3
                    onceexp = Val(temp1)
                    If onceexp = 0 Then MsgBox("这样的条件下是没有提督经验的哦(⊙o⊙)", vbInformation) : GoTo Input3
                Case 45
                    stage = "6-1N"
                    onceexp = 24
                    oncecost = 24
                Case 46
                    stage = "6-2N"
                    onceexp = 24
                    oncecost = 24
                Case 47
                    stage = "6-3N"
                    onceexp = 24
                    oncecost = 24
                Case 48
                    stage = "6-4N"
                    onceexp = 26
                    oncecost = 26
                Case 49
                    stage = "6-1H"
                    onceexp = 28
                    oncecost = 28
                Case 50
                    stage = "6-2H"
                    onceexp = 28
                    oncecost = 28
                Case 51
                    stage = "6-3H"
                    onceexp = 28
                    oncecost = 28
                Case 52
                    stage = "6-4H"
                    onceexp = 30
                    oncecost = 30
            End Select
            expfrom = tiduexp(CInt(textBoxfrom.Text))
            expto = tiduexp(CInt(textBoxto.Text))
            needexp = expto - expfrom
            times = needexp / onceexp
            elec = times * oncecost
            If level2 >= 40 And level2 <= 43 Then
                MsgBox("充满智慧的昆西掐指一算：" & "提督由等级" & textBoxfrom.Text & "升级至" & textBoxto.Text & "，共需经验值" & needexp & vbCrLf & "刷" & stage & "关卡的单个节点" & vbCrLf & "一次经验" & onceexp & "，需要次数：" & times & " 需要深海抑制装置：" & elec, vbInformation, "昆昆提示")
            ElseIf level2 = 44 Then
                MsgBox("充满智慧的昆西掐指一算：" & "提督由等级" & textBoxfrom.Text & "升级至" & textBoxto.Text & "，共需经验值" & needexp & vbCrLf & "刷" & stage & "关卡的单个节点" & vbCrLf & "一次经验" & onceexp & "，需要次数：" & times, vbInformation, "昆昆提示")
            Else
                MsgBox("充满智慧的昆西掐指一算：" & "提督由等级" & textBoxfrom.Text & "升级至" & textBoxto.Text & "，共需经验值" & needexp & vbCrLf & "刷" & stage & "关卡的单个节点" & vbCrLf & "一次经验" & onceexp & "，需要次数：" & times & " 需要电：" & elec, vbInformation, "昆昆提示")
            End If
        End If
    End Sub
End Class
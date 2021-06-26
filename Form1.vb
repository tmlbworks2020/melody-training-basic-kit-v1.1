'///////////////////////////////////
'Music_Melody Basic Kit ver1.1
' edited by tmlbworks (c)
' 2021/07 ver1.1 (VS2019) 
' 2020/06 ver1.0 (C++ Builder 1.03)
'///////////////////////////////////

Public Class Form1

    'CheckBoxコントロールの配列
    Private ArrCheckBox() As CheckBox
    Private ii As Integer
    Private mm As Integer

    'PLAY
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'melody check
        If MelodyLineChk() = True Then



            '各値読み込み
            Dim t2 As Integer = 50

            Dim t1 As Integer = CInt(60000 / CInt(ComboBox18.SelectedItem) - t2)

            Dim st As String = ""

            For j = 0 To 15
                st += textboxs(j).Text
            Next


            L5.Text = "1"
            L6.Text = comboboxs(0).SelectedItem

            '出音
            Dim n As Integer = 0
            'mm = 0
            Dim stp As Integer = CInt(ComboBox17.SelectedItem)
            For mm = 1 To stp
                L7.Text = CStr(mm)
                'Do While (True)
                '    If (mm = CInt(ComboBox17.SelectedItem)) Then Exit Do
                For ii = 0 To 15
                    '表示が1/4 遅れる
                    L5.Text = CStr(ii + 1)
                    L6.Text = comboboxs(ii).SelectedItem

                    For j = 1 To 7 Step 2
                        n = ii * 8 + j  '4だと 3/4 からで別パターン

                        Dim Nt As String = Strings.Mid(st, n, 2)

                        SoundOut(Nt, t1)
                        System.Threading.Thread.Sleep(t2)
                        Application.DoEvents()
                    Next

                Next
                '    mm += 1
                'Loop
            Next

            L5.Text = "1"
            L6.Text = comboboxs(0).SelectedItem

        Else
            MsgBox("Error! Melody Line の書式が不完全です")
        End If

    End Sub


    'Melody Text Check
    Private Function MelodyLineChk()
        Dim bl As Boolean = False
        Dim n As Integer = 0
        Dim m As Integer = 0
        Dim st As String = ""

        For j = 0 To 15
            st += textboxs(j).Text
        Next

        If st.Length <> 128 Then
            Return bl
        Else


            For ii = 0 To 15


                For j = 1 To 7 Step 2
                    n = ii * 8 + j  '4だと 3/4 からで別パターン

                    Dim Nt1 As String = Strings.Mid(st, n, 1)
                    Dim Nt2 As String = Strings.Mid(st, n + 1, 1)

                    If (Asc(Nt1) >= Asc("a") And Asc(Nt1) <= Asc("g")) Or (Asc(Nt1) >= Asc("A") And Asc(Nt1) <= Asc("G")) Then
                        If (Nt2 = "-" Or Nt2 = "#" Or Nt2 = "b") Then
                            bl = True
                        Else
                            Return False

                        End If

                    Else
                        Return False

                    End If


                Next

            Next
        End If

        Return bl
    End Function
    'PITCH LIST
    Private Sub SoundOut(ByRef Nt, ByRef tp)
        Dim p As Integer

        If (Nt = "c-") Then
            p = 523
        ElseIf (Nt = "c#" Or Nt = "db") Then
            p = 554
        ElseIf (Nt = "d-") Then
            p = 587
        ElseIf (Nt = "d#" Or Nt = "eb") Then
            p = 622
        ElseIf (Nt = "e-") Then
            p = 659
        ElseIf (Nt = "f-") Then
            p = 698
        ElseIf (Nt = "f#" Or Nt = "gb") Then

            p = 740
        ElseIf (Nt = "g-") Then

            p = 784
        ElseIf (Nt = "g#" Or Nt = "ab") Then

            p = 831
        ElseIf (Nt = "a-") Then

            p = 880
        ElseIf (Nt = "a#" Or Nt = "bb") Then

            p = 932
        ElseIf (Nt = "b-") Then

            p = 988
        ElseIf (Nt = "C-") Then

            p = 1047

        ElseIf (Nt = "C#" Or Nt = "Db") Then
            p = 1109

        ElseIf (Nt = "D-") Then
            p = 1175
        ElseIf (Nt = "D#" Or Nt = "Eb") Then
            p = 1245
        ElseIf (Nt = "E-") Then
            p = 1319
        ElseIf (Nt = "F-") Then
            p = 1397
        ElseIf (Nt = "F#" Or Nt = "Gb") Then

            p = 1480
        ElseIf (Nt = "G-") Then

            p = 1568
        ElseIf (Nt = "G#" Or Nt = "Ab") Then

            p = 1661

        ElseIf (Nt = "A-") Then

            p = 1760


        ElseIf (Nt = "A#" Or Nt = "Bb") Then

            p = 1865

        ElseIf (Nt = "B-") Then

            p = 1976
        End If


        Console.Beep(p, tp)


    End Sub

    'コントロール配列
    Private checkboxs() As System.Windows.Forms.CheckBox
    Private comboboxs() As System.Windows.Forms.ComboBox
    Private textboxs() As System.Windows.Forms.TextBox

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        '初期設定
        Me.checkboxs = New System.Windows.Forms.CheckBox(15) {}

        Me.checkboxs(0) = Me.CheckBox1
        Me.checkboxs(1) = Me.CheckBox2
        Me.checkboxs(2) = Me.CheckBox3
        Me.checkboxs(3) = Me.CheckBox4
        Me.checkboxs(4) = Me.CheckBox5
        Me.checkboxs(5) = Me.CheckBox6
        Me.checkboxs(6) = Me.CheckBox7
        Me.checkboxs(7) = Me.CheckBox8
        Me.checkboxs(8) = Me.CheckBox9
        Me.checkboxs(9) = Me.CheckBox10
        Me.checkboxs(10) = Me.CheckBox11
        Me.checkboxs(11) = Me.CheckBox12
        Me.checkboxs(12) = Me.CheckBox13
        Me.checkboxs(13) = Me.CheckBox14
        Me.checkboxs(14) = Me.CheckBox15
        Me.checkboxs(15) = Me.CheckBox16


        Me.comboboxs = New System.Windows.Forms.ComboBox(15) {}

        Me.comboboxs(0) = Me.ComboBox1
        Me.comboboxs(1) = Me.ComboBox2
        Me.comboboxs(2) = Me.ComboBox3
        Me.comboboxs(3) = Me.ComboBox4
        Me.comboboxs(4) = Me.ComboBox5
        Me.comboboxs(5) = Me.ComboBox6
        Me.comboboxs(6) = Me.ComboBox7
        Me.comboboxs(7) = Me.ComboBox8
        Me.comboboxs(8) = Me.ComboBox9
        Me.comboboxs(9) = Me.ComboBox10
        Me.comboboxs(10) = Me.ComboBox11
        Me.comboboxs(11) = Me.ComboBox12
        Me.comboboxs(12) = Me.ComboBox13
        Me.comboboxs(13) = Me.ComboBox14
        Me.comboboxs(14) = Me.ComboBox15
        Me.comboboxs(15) = Me.ComboBox16


        Me.textboxs = New System.Windows.Forms.TextBox(15) {}

        Me.textboxs(0) = Me.TextBox1
        Me.textboxs(1) = Me.TextBox2
        Me.textboxs(2) = Me.TextBox3
        Me.textboxs(3) = Me.TextBox4
        Me.textboxs(4) = Me.TextBox5
        Me.textboxs(5) = Me.TextBox6
        Me.textboxs(6) = Me.TextBox7
        Me.textboxs(7) = Me.TextBox8
        Me.textboxs(8) = Me.TextBox9
        Me.textboxs(9) = Me.TextBox10
        Me.textboxs(10) = Me.TextBox11
        Me.textboxs(11) = Me.TextBox12
        Me.textboxs(12) = Me.TextBox13
        Me.textboxs(13) = Me.TextBox14
        Me.textboxs(14) = Me.TextBox15
        Me.textboxs(15) = Me.TextBox16


        For i = 0 To 15
            checkboxs(i).CheckState = 1
        Next

        Dim aryCd() As String = {"C", "Cm", "C#", "C#m", "D", "Dm", "D#", "D#m", "E", "Em", "F", "Fm", "F#", "F#m", "G", "Gm", "G#", "G#m", "A", "Am", "A#", "A#m", "B", "Bm"}



        For j = 0 To 15
            comboboxs(j).Items.Clear()

            For Each i In aryCd
                comboboxs(j).Items.Add(i)
            Next

            comboboxs(j).SelectedIndex = 0
        Next


        'LOOP
        Dim aryLp() As String = {"1", "2", "3", "5", "10"}

        ComboBox17.Items.Clear()

        For Each i In aryLp
            ComboBox17.Items.Add(i)
        Next
        ComboBox17.SelectedIndex = 0

        'TEMPO
        Dim aryTp() As String = {"60", "80", "100", "120"}

        ComboBox18.Items.Clear()

        For Each i In aryTp
            ComboBox18.Items.Add(i)

        Next

        ComboBox18.SelectedIndex = 1

        'HALF NOTE
        Dim aryHf() As String = {"#", "b"}

        ComboBox19.Items.Clear()

        For Each i In aryHf
            ComboBox19.Items.Add(i)

        Next
        ComboBox19.SelectedIndex = 0

        'PATTERN
        Dim aryPt() As String = {"up", "down"} ', "up-down"

        ComboBox20.Items.Clear()

        For Each i In aryPt
            ComboBox20.Items.Add(i)
        Next
        ComboBox20.SelectedIndex = 0

    End Sub

    'MAKE
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        For j = 0 To 15
            If (comboboxs(j).SelectedIndex = -1) Then
                comboboxs(j).SelectedIndex = 0
            End If
        Next

        For i = 0 To 15
            If (checkboxs(i).CheckState = 1) Then

                textboxs(i).Text = ChordInput(comboboxs(i).SelectedIndex)
            End If
        Next
        'up, down ,ud,randomなど



    End Sub

    'CHORD COMPOSITION
    Function ChordInput(ByVal ch)
        Dim str As String = ""

        If (ComboBox19.SelectedIndex = 0) Then


            If (ComboBox20.SelectedIndex = 0) Then
                Select Case ch
                    Case 0 'c
                        str = "c-e-g-a#"
                    Case 1 'cm
                        str = "c-d#g-a#"
                    Case 2 'c#
                        str = "c#f-g#b-"
                    Case 3 'c#m
                        str = "c#e-g#b-"
                    Case 4 'd
                        str = "d-f#a-C-"
                    Case 5 'dm
                        str = "d-f-a-C-"
                    Case 6 'd#
                        str = "d#g-a#C#"
                    Case 7 'd#m
                        str = "d#f#a#C#"
                    Case 8 'e
                        str = "e-g#b-D-"
                    Case 9 'em
                        str = "e-g-b-D-"
                    Case 10 'f
                        str = "f-F-f-F-"
                    Case 11 'fm
                        str = "f-g#C-D#"
                    Case 12 'f#
                        str = "f#a#C#E-"
                    Case 13 'f#m
                        str = "f#a-C#E-"
                    Case 14 'g
                        str = "g-b-D-F-"
                    Case 15 'gm
                        str = "g-a#D-F-"
                    Case 16 'g#
                        str = "g#C-D#F#"
                    Case 17 'g#m
                        str = "g#b-D#F#"
                    Case 18 'a
                        str = "a-C#E-G-"
                    Case 19 'am
                        str = "a-C-E-G-"
                    Case 20 'a#
                        str = "a#D-F-G#"
                    Case 21 'a#m
                        str = "a#C#F-G#"
                    Case 22 'b
                        str = "b-D#F#A-"
                    Case 23 'bm
                        str = "b-D-F#A-"
                End Select

            ElseIf (ComboBox20.SelectedIndex = 1) Then
                Select Case ch
                    Case 0 'c
                        str = "C-a#g-e-"
                    Case 1 'cm
                        str = "C-a#g-d#"
                    Case 2 'c#
                        str = "C#b-g#f-"
                    Case 3 'c#m
                        str = "C#b-g#e-"
                    Case 4 'd
                        str = "D-C-a-f#"
                    Case 5 'dm
                        str = "D-C-a-f-"
                    Case 6 'd#
                        str = "D#C#a#g-"
                    Case 7 'd#m
                        str = "D#C#a#f#"
                    Case 8 'e
                        str = "E-D-b-g#"
                    Case 9 'em
                        str = "E-D-b-g-"
                    Case 10 'f
                        str = "F-D#C-a-"
                    Case 11 'fm
                        str = "F-D#C-g#"
                    Case 12 'f#
                        str = "F#E-C#a#"
                    Case 13 'f#m
                        str = "F#E-C#a-"
                    Case 14 'g-
                        str = "G-g-G-g-"
                    Case 15 'gm
                        str = "G-F-D-a#"
                    Case 16 'g#
                        str = "g#f#d#c-"
                    Case 17 'g#m
                        str = "g#f#d#b-"
                    Case 18 'a
                        str = "a-g-e-c#"
                    Case 19 'am
                        str = "a-g-e-c-"
                    Case 20 'a#
                        str = "a#g#f-d-"
                    Case 21 'a#m
                        str = "a#g#f-c#"
                    Case 22 'b
                        str = "b-a-f#d#"
                    Case 23 'bm
                        str = "b-a-f#d-"
                End Select
            End If

        ElseIf (ComboBox19.SelectedIndex = 1) Then

            If (ComboBox20.SelectedIndex = 0) Then
                Select Case ch
                    Case 0 'c
                        str = "c-e-g-bb"
                    Case 1 'cm
                        str = "c-ebg-bb"
                    Case 2 'db
                        str = "dbf-abb-"
                    Case 3 'dbm
                        str = "dbe-abb-"
                    Case 4 'd
                        str = "d-gba-C-"
                    Case 5 'dm
                        str = "d-f-a-C-"
                    Case 6 'eb
                        str = "ebg-bbDb"
                    Case 7 'ebm
                        str = "ebgbbbDb"
                    Case 8 'e
                        str = "e-abb-D-"
                    Case 9 'em
                        str = "e-g-b-D-"
                    Case 10 'f
                        str = "f-a-C-Eb"
                    Case 11 'fm
                        str = "f-abC-Eb"
                    Case 12 'gb
                        str = "gbbbDbE-"
                    Case 13 'gbm
                        str = "gba-DbE-"
                    Case 14 'g
                        str = "g-b-D-F-"
                    Case 15 'gm
                        str = "g-bbD-F-"
                    Case 16 'ab
                        str = "abC-EbGb"
                    Case 17 'abm
                        str = "abb-EbGb"
                    Case 18 'a
                        str = "a-DbE-G-"
                    Case 19 'am
                        str = "a-C-E-G-"
                    Case 20 'bb
                        str = "bbD-F-Ab"
                    Case 21 'bbm
                        str = "bbDbF-Ab"
                    Case 22 'b
                        str = "b-EbGbA-"
                    Case 23 'bm
                        str = "b-D-GbA-"
                End Select

            ElseIf (ComboBox20.SelectedIndex = 1) Then
                Select Case ch
                    Case 0 'c
                        str = "C-bbg-e-"
                    Case 1 'cm
                        str = "C-bbg-eb"
                    Case 2 'db
                        str = "Dbb-abf-"
                    Case 3 'dbm
                        str = "Dbb-abe-"
                    Case 4 'd
                        str = "D-C-a-gb"
                    Case 5 'dm
                        str = "D-C-a-f-"
                    Case 6 'eb
                        str = "EbDbbbg-"
                    Case 7 'ebm
                        str = "EbDbbbgb"
                    Case 8 'e
                        str = "E-D-b-ab"
                    Case 9 'em
                        str = "E-D-b-g-"
                    Case 10 'f
                        str = "F-EbC-a-"
                    Case 11 'fm
                        str = "F-EbC-ab"
                    Case 12 'gb
                        str = "GbE-Dbbb"
                    Case 13 'gbm
                        str = "GbE-Dba-"
                    Case 14 'g-
                        str = "G-F-D-b-"
                    Case 15 'gm
                        str = "G-F-D-bb"
                    Case 16 'ab
                        str = "abgbebc-"
                    Case 17 'abm
                        str = "abgbebb-"
                    Case 18 'a
                        str = "a-g-e-db"
                    Case 19 'am
                        str = "a-g-e-c-"
                    Case 20 'bb
                        str = "bbabf-d-"
                    Case 21 'bbm
                        str = "bbabf-db"
                    Case 22 'b
                        str = "b-a-gbeb"
                    Case 23 'bm
                        str = "b-a-gbd-"
                End Select
            End If

        End If




        Return str
    End Function


    'STOP 16 BARS
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        Select Case ComboBox17.SelectedIndex
            'Case 0
            '    mm = 2
            'Case 1
            '    mm = 3
            'Case 2
            '    mm = 4
            'Case 3
            '    mm = 6
            'Case 4
            '    mm = 11
            Case 0
                mm = 1
            Case 1
                mm = 2
            Case 2
                mm = 3
            Case 3
                mm = 5
            Case 4
                mm = 10
        End Select

        ii = 15


    End Sub


    'Exit
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click


        Select Case ComboBox17.SelectedIndex
            Case 0
                mm = 1
            Case 1
                mm = 2
            Case 2
                mm = 3
            Case 3
                mm = 5
            Case 4
                mm = 10
                'Case 0
                '    mm = 2
                'Case 1
                '    mm = 3
                'Case 2
                '    mm = 4
                'Case 3
                '    mm = 6
                'Case 4
                '    mm = 11
        End Select
        ii = 15
        Me.Close()
    End Sub

    'OPEN
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Using ofd As OpenFileDialog = New OpenFileDialog
            'フォルダパスを exe のフォルダにする
            ofd.InitialDirectory = System.Environment.CurrentDirectory
            'デフォルトのファイル名を指定します
            ofd.FileName = "melody.dat"

            '選択できるファイルの種類（拡張子）を限定します
            ofd.Filter = "dat |*.dat" ' CSV Files |*.csv"

            If ofd.ShowDialog() = DialogResult.OK Then
                Using sr As New System.IO.StreamReader(ofd.FileName, System.Text.Encoding.GetEncoding("Shift-JIS"))
                    'Dim line As String = ""
                    Do

                        '１行ずつファイルを読み込み
                        'line = sr.ReadLine()
                        ComboBox17.SelectedIndex = sr.ReadLine()
                        ComboBox18.SelectedIndex = sr.ReadLine()
                        ComboBox19.SelectedIndex = sr.ReadLine()
                        ComboBox20.SelectedIndex = sr.ReadLine()

                        For i = 0 To 15
                            checkboxs(i).CheckState = sr.ReadLine()
                            comboboxs(i).SelectedIndex = sr.ReadLine()
                            textboxs(i).Text = sr.ReadLine()

                        Next
                        'Console.WriteLine(line)
                    Loop Until sr.ReadLine() Is Nothing
                End Using
            End If
            TextBox17.Text = ofd.FileName
        End Using

    End Sub


    'SAVE Dialog
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        'SaveFileDialogクラスのインスタンスを作成
        Dim sfd As New SaveFileDialog()

        sfd.InitialDirectory = System.Environment.CurrentDirectory
        'sfd.InitialDirectory = "."
        sfd.FileName = "melody.dat"


        sfd.Filter = "datファイル(*.dat)|*.dat"
        '2番目の「すべてのファイル」が選択されているようにする
        sfd.FilterIndex = 2
        sfd.Title = "保存先のファイルを選択してください"
        'ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
        sfd.RestoreDirectory = True
        '既に存在するファイル名を指定したとき警告する
        sfd.OverwritePrompt = True
        '存在しないパスが指定されたとき警告を表示する
        sfd.CheckPathExists = True

        'ダイアログを表示する
        If sfd.ShowDialog() = DialogResult.OK Then
            'OKボタンがクリックされたとき、選択されたファイル名を表示する

            'Console.WriteLine(sfd.FileName)
            SaveFile(sfd.FileName)
            TextBox17.Text = sfd.FileName
        End If
    End Sub

    'Save
    Private Sub SaveFile(ByVal fn)

        Dim str As String = ""

        Dim Writer As IO.StreamWriter
        Dim Encode As System.Text.Encoding
        '文字コードにShiftJISを指定。(UTF8の場合は指定不要)
        Encode = System.Text.Encoding.GetEncoding("Shift-JIS")

        '既に存在するテキストに追加する場合は第２引数をTrueにする。
        Writer = New IO.StreamWriter(fn, False, Encode)



        Writer.WriteLine(ComboBox17.SelectedIndex)
        Writer.WriteLine(ComboBox18.SelectedIndex)
        Writer.WriteLine(ComboBox19.SelectedIndex)
        Writer.WriteLine(ComboBox20.SelectedIndex)

        For i = 0 To 15
            Writer.WriteLine(checkboxs(i).CheckState)
            Writer.WriteLine(comboboxs(i).SelectedIndex)
            Writer.WriteLine(textboxs(i).Text)
        Next


        Writer.Close()



    End Sub

    'b-# change
    Private Sub ComboBox19_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox19.SelectedIndexChanged

        Dim k As Integer = 0


        Dim aryCd() As String = {}
        'Chord change
        If (ComboBox19.SelectedIndex = 1) Then
            aryCd = {"C", "Cm", "Db", "Dbm", "D", "Dm", "Eb", "Ebm", "E", "Em", "F", "Fm", "Gb", "Gbm", "G", "Gm", "Ab", "Abm", "A", "Am", "Bb", "Bbm", "B", "Bm"}

            'melody change

            For k = 0 To 15
                Dim mel = textboxs(k).Text

                If (mel.Length > 0) Then
                    Dim arym() As String = {"", "", "", "", "", "", "", ""}

                    For j = 1 To 7 Step 2
                        Dim Nt As String = Strings.Mid(mel, j + 1, 1)

                        If (Nt = "#") Then
                            arym(j) = "b"
                            Dim c1 As Char = Mid(mel, j, 1)
                            If (c1 = "g") Then
                                arym(j - 1) = "a"
                            ElseIf (c1 = "G") Then
                                arym(j - 1) = "A"
                            Else
                                arym(j - 1) = CStr(Chr(Asc(c1) + 1))
                            End If

                        Else
                            arym(j - 1) = Strings.Mid(mel, j, 1)
                            arym(j) = Nt

                        End If

                    Next


                    textboxs(k).Clear()
                    For i = 0 To 7
                        textboxs(k).Text += arym(i)
                    Next

                End If
            Next



        ElseIf (ComboBox19.SelectedIndex = 0) Then
            aryCd = {"C", "Cm", "C#", "C#m", "D", "Dm", "D#", "D#m", "E", "Em", "F", "Fm", "F#", "F#m", "G", "Gm", "G#", "G#m", "A", "Am", "A#", "A#m", "B", "Bm"}

            'melody change

            For k = 0 To 15
                Dim mel = textboxs(k).Text

                If (mel.Length > 0) Then
                    Dim arym() As String = {"", "", "", "", "", "", "", ""}

                    For j = 1 To 7 Step 2
                        Dim Nt As String = Strings.Mid(mel, j + 1, 1)

                        If (Nt = "b") Then
                            arym(j) = "#"
                            Dim c1 As Char = Mid(mel, j, 1)
                            If (c1 = "a") Then
                                arym(j - 1) = "g"
                            ElseIf (c1 = "A") Then
                                arym(j - 1) = "G"
                            Else
                                arym(j - 1) = CStr(Chr(Asc(c1) - 1))
                            End If

                        Else
                            arym(j - 1) = Strings.Mid(mel, j, 1)
                            arym(j) = Nt

                        End If

                    Next


                    textboxs(k).Clear()
                    For i = 0 To 7
                        textboxs(k).Text += arym(i)
                    Next

                End If



            Next



        End If


        For j = 0 To 15

            k = comboboxs(j).SelectedIndex
            comboboxs(j).Items.Clear()

            For Each i In aryCd
                comboboxs(j).Items.Add(i)
            Next

            comboboxs(j).SelectedIndex = k
        Next


    End Sub
End Class

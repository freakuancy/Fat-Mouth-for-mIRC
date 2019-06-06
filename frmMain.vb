


Imports System
Imports System.Drawing.Design
Imports System.IO
Imports System.Collections
Imports System.Text.RegularExpressions
Imports FatMouth.Org.Mentalis.Files


Public Class frmMain
    Public LastChan As String = Nothing
    Public cntDialog As Integer
    Public cntForeign As Integer
    Public cntTelepathy As Integer
    Public cntParagraphs As Integer
    Public sndCount As Integer
    Public Sub DoQuit()

        With My.Settings
            '.colDefault = ComboBox1.ComboBox.SelectedIndex
            .colDialog = ComboBox3.SelectedIndex
            .colTelepathy = ComboBox4.SelectedIndex
            .colForeign = ComboBox5.SelectedIndex
            .colDone = ComboBox6.SelectedIndex
            .colNext = ComboBox7.SelectedIndex
            .colContinue = ComboBox8.SelectedIndex
            .colBegin = ComboBox9.SelectedIndex
            .txtDialog = TextBox1.Text
            .txtTelepathy = TextBox2.Text
            .txtForeign = TextBox3.Text
            .txtDone = TextBox4.Text
            .txtNext = TextBox5.Text
            .txtContinue = TextBox6.Text
            .txtBegin = TextBox7.Text
            .AlwaysTop = Me.TopMost
            .CutPreview = CheckBox2.Checked
            .AutoColor = CheckBox1.Checked
            .ClearSend = ToolStripMenuItem32.Checked
            .Save()
        End With
    End Sub
    Public Sub LoadProfile(ByVal pName As String)

        Dim iniSettings As New IniReader(pName)
        Dim TempTable(15) As String

        ComboBox1.Items.Clear()
        ComboBox3.Items.Clear()
        ComboBox4.Items.Clear()
        ComboBox5.Items.Clear()
        ComboBox6.Items.Clear()
        ComboBox7.Items.Clear()
        ComboBox8.Items.Clear()
        ComboBox9.Items.Clear()

        For i = 0 To 15
            ComboBox1.ComboBox.Items.Add("Color " + i.ToString("00"))
            ComboBox3.Items.Add("Color " + i.ToString("00"))
            ComboBox4.Items.Add("Color " + i.ToString("00"))
            ComboBox5.Items.Add("Color " + i.ToString("00"))
            ComboBox6.Items.Add("Color " + i.ToString("00"))
            ComboBox7.Items.Add("Color " + i.ToString("00"))
            ComboBox8.Items.Add("Color " + i.ToString("00"))
            ComboBox9.Items.Add("Color " + i.ToString("00"))
            If My.Settings.Standalone = False Then
                TempTable(i) = Integer.Parse(iniSettings.ReadString("Palette", i.ToString("00")))
                Dim Scolor As Color = Color.FromArgb(255, Color.FromArgb(TempTable(i)))
                Dim cRed As Byte = Scolor.R
                Dim cBlue As Byte = Scolor.B
                Dim cGreen As Byte = Scolor.G
                colTable(i) = Color.FromArgb(cBlue, cGreen, cRed)
            End If
        Next
        'Fill Elements structure
        If My.Settings.Standalone = False Then
            colElements._colDefault = iniSettings.ReadString("Colors", "colDefault")
            colElements._colAction = iniSettings.ReadString("Colors", "colAction")
            colElements._colDialog = iniSettings.ReadString("Colors", "colDialog")
            colElements._colTelepathy = iniSettings.ReadString("Colors", "colTelepathy")
            colElements._colForeign = iniSettings.ReadString("Colors", "colForeign")
            colElements._colNext = iniSettings.ReadString("Colors", "colNext")
            colElements._colBegin = iniSettings.ReadString("Colors", "colBegin")
            colElements._colCont = iniSettings.ReadString("Colors", "colContinue")
            colElements._colDefault = iniSettings.ReadString("Colors", "colDefault")
            colElements._colDefault = iniSettings.ReadString("Colors", "colDefault")
            colElements._colDone = iniSettings.ReadString("Colors", "colDone")
            colElements._txtBegin = iniSettings.ReadString("Brackets", "txtBegin")
            colElements._txtDone = iniSettings.ReadString("Brackets", "txtDone")
            colElements._txtCont = iniSettings.ReadString("Brackets", "txtContinue")
            colElements._txtNext = iniSettings.ReadString("Brackets", "txtNext")
            colElements._txtDialog = iniSettings.ReadString("Chars", "txtDialog")
            colElements._txtForeign = iniSettings.ReadString("Chars", "txtForeign")
            colElements._txtTelepathy = iniSettings.ReadString("Chars", "txtTelepathy")
        End If
        ComboBox1.SelectedIndex = colElements._colDefault
        ComboBox3.SelectedIndex = colElements._colDialog
        ComboBox4.SelectedIndex = colElements._colTelepathy
        ComboBox5.SelectedIndex = colElements._colForeign
        ComboBox6.SelectedIndex = colElements._colDone
        ComboBox7.SelectedIndex = colElements._colNext
        ComboBox8.SelectedIndex = colElements._colCont
        ComboBox9.SelectedIndex = colElements._colBegin
        TextBox1.Text = colElements._txtDialog
        TextBox2.Text = colElements._txtTelepathy
        TextBox3.Text = colElements._txtForeign
        TextBox4.Text = colElements._txtDone
        TextBox5.Text = colElements._txtNext
        TextBox6.Text = colElements._txtCont
        TextBox7.Text = colElements._txtBegin

        rchBoxMain.ForeColor = colTable(colElements._colDefault)
        rchBoxMain.SelectionColor = colTable(colElements._colDefault)
        If CheckBox1.Checked = True Then
            With rchBoxMain
                Dim oldpos As Integer = .SelectionStart

                If CheckBox1.Checked = True Then

                    If rchBoxMain.TextLength <> 0 Then
                        SuspendDrawing(rchBoxMain)
                        .SelectionStart = 1
                        .SelectionLength = .Text.Length
                        .SelectionColor = .ForeColor
                        .SelectionLength = 0
                        .SelectionStart = oldpos
                        MatchQuotes()
                        MatchForeign()
                        MatchTele()
                        .SelectionLength = 0
                        .SelectionStart = oldpos
                    End If
                End If
                SuspendDrawing(rchBoxMain)
                .SelectionStart = 1
                .SelectionLength = .Text.Length
                .SelectionBackColor = .BackColor
                MatchLength()
                .SelectionLength = 0
                .SelectionStart = oldpos


            End With
            ResumeDrawing(rchBoxMain, True)
        End If
    End Sub
    Public Function ParseText(Optional ByVal Segment As Integer = 0)
        Dim ch As String
        Dim fString As String = vbNullString
        Dim isBold As Boolean
        Dim isItalic As Boolean
        Dim isUline As Boolean
        Dim cCode As Color
        Dim colCode As String = vbNullString
        Dim rchIndex As Integer = 0
        Dim indCount As Integer = 0
        If CheckBox3.Checked = True Then
            Select Case Segment
                Case 0
                    If colElements._txtBegin.Length <> 0 Then
                        fString = Chr(3) + ComboBox9.SelectedIndex.ToString("00") + TextBox7.Text + Chr(32)
                    End If
                Case 1, 2
                    If colElements._txtCont.Length <> 0 Then
                        fString = Chr(3) + ComboBox8.SelectedIndex.ToString("00") + TextBox6.Text + Chr(32)
                    End If
            End Select
        End If
        For Each ch In RichTextBox1.Text
            With RichTextBox1
                .SelectionStart = rchIndex
                .SelectionLength = 1
                If .SelectionFont.Bold Then
                    If isBold = False Then
                        fString = fString + Chr(2)
                        isBold = True
                    End If
                Else
                    If isBold = True Then
                        isBold = False
                        fString = fString + Chr(2)
                    End If
                End If
                If .SelectionFont.Underline Then
                    If isUline = False Then
                        fString = fString + Chr(31)
                        isUline = True
                    End If
                Else
                    If isUline = True Then
                        isUline = False
                        fString = fString + Chr(31)
                    End If
                End If

                If .SelectionFont.Italic Then
                    If isItalic = False Then
                        fString = fString + Chr(29)
                        isItalic = True
                    End If
                Else
                    If isItalic = True Then
                        isItalic = False
                        fString = fString + Chr(29)
                    End If
                End If
                If .SelectionColor <> cCode Then
                    indCount = 0
                    cCode = .SelectionColor
                    For Each cv In colTable

                        Dim nColor As Color = Color.FromArgb(255, cv)
                        If nColor.ToArgb.ToString("x") = cCode.ToArgb.ToString("x") Then
                            colCode = indCount.ToString("00")
                            Exit For
                        End If
                        indCount = indCount + 1
                    Next
                    fString = fString + Chr(3) + colCode
                End If

            End With
            fString = fString + ch
            rchIndex = rchIndex + 1

        Next
        If CheckBox3.Checked = True Then
            Select Case Segment
                Case 0, 1
                    If colElements._txtNext.Length <> 0 Then
                        fString = Trim(fString) + Chr(32) + Chr(3) + ComboBox7.SelectedIndex.ToString("00") + TextBox5.Text
                    End If
                Case 2
                    If colElements._txtDone.Length <> 0 Then
                        fString = Trim(fString) + Chr(32) + Chr(3) + ComboBox6.SelectedIndex.ToString("00") + TextBox4.Text
                    End If
            End Select
        End If
        Return fString
    End Function
    Public Sub MatchLength()
        Dim reges As String = "(^.*\s)"

        Dim options As RegexOptions = RegexOptions.Multiline
        Dim reg As New System.Text.RegularExpressions.Regex(reges, options)
        Dim src As String = Me.rchBoxMain.Text + Chr(32)
        Dim mat As System.Text.RegularExpressions.Match = reg.Match(src)
        Dim mats As System.Text.RegularExpressions.MatchCollection = reg.Matches(src)
        Dim curPosition As Integer = rchBoxMain.SelectionStart
        Dim Backlight As Color = Color.FromArgb(1, Color.DarkGray)
        Dim oldColor As Color = rchBoxMain.BackColor
        sndCount = 0


        With rchBoxMain

            While mat.Success
                Dim selCount As Integer = 0

                If mat.Length > 370 Then
                    sndCount = sndCount + (Math.Ceiling(mat.Length / 370))
                    If CheckBox2.Checked = True Then
                        .SelectionStart = mat.Index + 370
                        .SelectionLength = 1

                        'Search backwards for nearest space, up to 64 characters
                        Do While Not .SelectedText = Chr(32) And Not selCount = 64
                            .SelectionStart = (.SelectionStart - 1)
                            selCount = selCount + 1
                        Loop

                        .SelectionStart = .SelectionStart + 1
                        .SelectionLength = (mat.Length - (370 - (selCount - 2)))
                        .SelectionBackColor = Backlight
                    End If
                Else
                    If mat.Length <> 1 Then sndCount = sndCount + 1
                End If

                mat = reg.Match(src, mat.Index + (mat.Length))
                ToolStripProgressBar1.Maximum = sndCount + 1
                lblSends.Text = "Sends: " + sndCount.ToString("00")
            End While

            .SelectionStart = curPosition
        End With

    End Sub

    Public Sub MatchQuotes()
        Dim regexes As String = "(" + Regex.Escape(TextBox1.Text) + "(.*?)" + Regex.Escape(TextBox1.Text) + ")"
        Dim reg As New System.Text.RegularExpressions.Regex(regexes)
        Dim mat As System.Text.RegularExpressions.Match = reg.Match(Me.rchBoxMain.Text)
        Dim mats As System.Text.RegularExpressions.MatchCollection = reg.Matches(Me.rchBoxMain.Text)
        Dim curPosition As Integer = rchBoxMain.SelectionStart
        Dim oldColor As Color = rchBoxMain.SelectionColor
        With rchBoxMain

            While mat.Success
                .SelectionStart = mat.Index
                .SelectionLength = mat.Length
                .SelectionColor = colTable(ComboBox3.SelectedIndex)
                mat = reg.Match(.Text, mat.Index + (mat.Length))
            End While

            .SelectionStart = curPosition
            .SelectionColor = oldColor
        End With

    End Sub
    Public Sub MatchTele()
        Dim reges As String = "(" + Regex.Escape(TextBox2.Text) + "(.*?)" + Regex.Escape(TextBox2.Text) + ")"
        Dim reg As New System.Text.RegularExpressions.Regex(reges)
        Dim mat As System.Text.RegularExpressions.Match = reg.Match(Me.rchBoxMain.Text)
        Dim mats As System.Text.RegularExpressions.MatchCollection = reg.Matches(Me.rchBoxMain.Text)

        Dim curPosition As Integer = rchBoxMain.SelectionStart
        Dim oldColor As Color = rchBoxMain.SelectionColor

        With rchBoxMain

            While mat.Success
                .SelectionStart = mat.Index
                .SelectionLength = mat.Length
                .SelectionColor = colTable(ComboBox4.SelectedIndex)
                mat = reg.Match(.Text, mat.Index + (mat.Length))
            End While

            .SelectionStart = curPosition
            .SelectionColor = oldColor
        End With

    End Sub
    Public Sub MatchForeign()
        Dim reges As String = "(" + Regex.Escape(TextBox3.Text) + "(.*?)" + Regex.Escape(TextBox3.Text) + ")"
        Dim reg As New System.Text.RegularExpressions.Regex(reges)
        Dim mat As System.Text.RegularExpressions.Match = reg.Match(Me.rchBoxMain.Text)
        Dim mats As System.Text.RegularExpressions.MatchCollection = reg.Matches(Me.rchBoxMain.Text)
        Dim curPosition As Integer = rchBoxMain.SelectionStart
        Dim oldColor As Color = rchBoxMain.SelectionColor
        With rchBoxMain

            While mat.Success
                .SelectionStart = mat.Index
                .SelectionLength = mat.Length
                .SelectionColor = colTable(ComboBox5.SelectedIndex)
                mat = reg.Match(.Text, mat.Index + (mat.Length))
            End While

            .SelectionStart = curPosition
            .SelectionColor = oldColor
        End With

    End Sub


    Private Sub ToolStripDropDownButton1_ButtonClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripDropDownButton1.ButtonClick

        Dim index = 0
        Dim src = rchBoxMain.Text + Chr(32)
        Dim srcNew As String
        Dim maxIndex As Integer = src.Length
        Dim cnames() As String = GetChannelNames()
        Dim chandles() As Integer = GetChannelHandle()
        Dim results As New List(Of String)
        Dim mirc As Integer
        Dim mdiclient As Integer
        Dim cntSeg As Integer = 1
        Dim mircchannel As Integer
        Dim Segment As Integer
        Dim cnt As Integer = 0
        If rchBoxMain.TextLength = 0 Then Exit Sub

        mirc = FindWindow("mirc", vbNullString)
        mdiclient = FindWindowEx(mirc, 0&, "mdiclient", vbNullString)
        mircchannel = FindWindowEx(mdiclient, 0&, "mirc_channel", vbNullString)

        If LastChan <> Nothing And My.Settings.RememberChannel = True Then
            cnt = Array.IndexOf(cnames, LastChan)
            mircchannel = chandles(cnt)
        End If

        ToolStripProgressBar1.Visible = True

        If rchBoxMain.TextLength >= 370 Then
            ToolStripProgressBar1.Increment(1)

            Do While index <= maxIndex

                Select Case cntSeg
                    Case 1 : Segment = 0
                    Case sndCount : Segment = 2
                    Case Else : Segment = 1
                End Select

                srcNew = src.Substring(index)
                Dim match = Regex.Match(srcNew, "^.{0," & 370 & "}\s").ToString()
                rchBoxMain.SelectionStart = index
                rchBoxMain.SelectionLength = match.Length

                If match.Length <= 1 Then
                    If cntSeg > sndCount Or match.Length = 0 Then
                        ToolStripProgressBar1.Value = 1
                        ToolStripProgressBar1.Visible = False
                        Exit Sub
                    End If

                    ToolStripProgressBar1.Increment(1)
                    results.Add(match)
                    index = index + match.Length
                    Continue Do
                End If

                RichTextBox1.Rtf = rchBoxMain.SelectedRtf
                ToolStripProgressBar1.Increment(1)
                mircSendText(ParseText(Segment), mircchannel)
                cntSeg = cntSeg + 1
                Threading.Thread.Sleep(My.Settings.SendDelay * 100)
                results.Add(match)
                index = index + match.Length

            Loop
        Else
            RichTextBox1.Rtf = rchBoxMain.Rtf
            mircSendText(ParseText(), mircchannel)
        End If
        LastChan = cnames(cnt)
        ToolStripProgressBar1.Visible = False
        ToolStripProgressBar1.Value = 1
        If ToolStripMenuItem32.Checked = True Then Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub rchBoxMain_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rchBoxMain.Click
        Dim cCode As Color = rchBoxMain.SelectionColor
        Dim i As Integer = 0
        For Each cv As Color In colTable
            If cv.ToArgb.ToString("x") = cCode.ToArgb.ToString("x") Then
                ComboBox1.ComboBox.SelectedIndex = i
            End If
            i = i + 1
        Next
    End Sub



    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        With rchBoxMain
            If .SelectionFont.Bold Then
                .SelectionFont = New Font(.SelectionFont, _
                                          .SelectionFont.Style Xor FontStyle.Bold)
            Else
                .SelectionFont = New Font(.SelectionFont, _
                                          .SelectionFont.Style Or FontStyle.Bold)
            End If
        End With
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        With rchBoxMain
            If .SelectionFont.Bold Then
                .SelectionFont = New Font(.SelectionFont, _
                                          .SelectionFont.Style Xor FontStyle.Italic)
            Else
                .SelectionFont = New Font(.SelectionFont, _
                                          .SelectionFont.Style Or FontStyle.Italic)
            End If
        End With
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        With rchBoxMain
            If .SelectionFont.Bold Then
                .SelectionFont = New Font(.SelectionFont, _
                                          .SelectionFont.Style Xor FontStyle.Underline)
            Else
                .SelectionFont = New Font(.SelectionFont, _
                                          .SelectionFont.Style Or FontStyle.Underline)
            End If
        End With
    End Sub

    Private Sub frmMain_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Dim cnames As String() = GetChannelNames()
        lblChannel.Text = "Active Channel: " + cnames(0)
    End Sub

    Private Sub frmMain_UnLoad(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Disposed
        DoQuit()

    End Sub
    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        With My.Settings

            If .FirstRun = True Then
                Me.TopMost = False
                FMOptions.Show()
                FMOptions.TopMost = True
                Me.Hide()
            Else
                Me.TopMost = .AlwaysTop
                FillForm()
            End If

            ComboBox1.ComboBox.DrawMode = DrawMode.OwnerDrawVariable
            AddHandler ComboBox1.ComboBox.DrawItem, AddressOf DrawItem
            AddHandler ComboBox3.DrawItem, AddressOf DrawItem
            AddHandler ComboBox4.DrawItem, AddressOf DrawItem
            AddHandler ComboBox5.DrawItem, AddressOf DrawItem
            AddHandler ComboBox6.DrawItem, AddressOf DrawItem
            AddHandler ComboBox7.DrawItem, AddressOf DrawItem
            AddHandler ComboBox8.DrawItem, AddressOf DrawItem
            AddHandler ComboBox9.DrawItem, AddressOf DrawItem

        End With
        rchBoxMain.Clear()
    End Sub


    Private Sub DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs)
        'Custom Color-Picker DropBox renders mIRC Palette
        Dim drawFont As Font = e.Font
        If e.Index = -1 Then Return
        Dim textBrush As New SolidBrush(Color.FromArgb(255, colTable(e.Index)))

        If (e.State And DrawItemState.Selected) = DrawItemState.Selected Then
            e.Graphics.FillRectangle(Brushes.WhiteSmoke, e.Bounds)
        End If

        e.DrawBackground()
        e.Graphics.DrawRectangle(Pens.Black, 3, e.Bounds.Top + 1, 13, 11)
        e.Graphics.FillRectangle(textBrush, 4, e.Bounds.Top + 2, 12, 10)

        textBrush.Dispose()
        e.DrawFocusRectangle()
    End Sub


    Private Sub rchBoxMain_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rchBoxMain.TextChanged

        With rchBoxMain
            Dim oldpos As Integer = .SelectionStart

            If CheckBox1.Checked Then
                If rchBoxMain.TextLength <> 0 Then 'Do Auto-Color
                    SuspendDrawing(rchBoxMain) 'Kill control paints until update is completed
                    .SelectionStart = 1
                    .SelectionLength = .Text.Length
                    .SelectionColor = .ForeColor
                    .SelectionLength = 0
                    .SelectionStart = oldpos
                    MatchQuotes()
                    MatchForeign()
                    MatchTele()
                    .SelectionLength = 0
                    .SelectionStart = oldpos
                End If
            End If
            SuspendDrawing(rchBoxMain)
            .SelectionStart = 1
            .SelectionLength = .Text.Length
            .SelectionBackColor = .BackColor
            MatchLength() 'Cut Highlights and Send Count
            .SelectionLength = 0
            .SelectionStart = oldpos

        End With

        ResumeDrawing(rchBoxMain, True) 'Update completed, redraw

        ' Count words
        Dim WordCounter As Integer = UBound(Split(Trim(Replace(rchBoxMain.Text, Space(2), Space(1))))) + 1
        If rchBoxMain.TextLength > 0 Then
            lblWords.Text = "Words: " + WordCounter.ToString
        Else
            lblWords.Text = "Words: " + "0"
        End If

    End Sub

    Private Sub cmbProfiles_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbProfiles.SelectedIndexChanged
        Dim pName As String

        pName = My.Settings.VPath + "\" + cmbProfiles.Text + ".clr"
        LoadProfile(pName)
    End Sub

    Private Sub FormatToolbarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FormatToolbarToolStripMenuItem.Click

        If ToolStrip1.Visible Then
            ToolStrip1.Hide()
        Else
            ToolStrip1.Show()
        End If
    End Sub

    Private Sub OptionsPanelToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptionsPanelToolStripMenuItem.Click

        If Panel1.Visible Then
            Panel1.Hide()
            rchBoxMain.Left = MenuStrip2.Left + 2
            rchBoxMain.Width = Me.Width + 2
        Else
            Panel1.Show()
            rchBoxMain.Left = Panel1.Right + 2
            rchBoxMain.Width = rchBoxMain.Width - Panel1.Width - 2
        End If

    End Sub

    Private Sub StatusBarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StatusBarToolStripMenuItem.Click

        If StatusStrip1.Visible Then
            StatusStrip1.Hide()
        Else
            StatusStrip1.Show()
        End If

    End Sub

    Private Sub ToolStripMenuItem10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem10.Click

        OpenFileDialog1.Filter = "Richtext Files (*.rtf)|*.rtf|Text Files (*.txt)|*.txt"
        OpenFileDialog1.Title = "Select a Rich-Text Formatted File to Open..."
        OpenFileDialog1.FileName = "*.rtf"
        OpenFileDialog1.ShowDialog()

        Dim sFile As String = OpenFileDialog1.FileName

        If File.Exists(sFile) Then rchBoxMain.LoadFile(sFile)

    End Sub

    Private Sub ToolStripMenuItem11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem11.Click

        SaveFileDialog1.Filter = "Richtext Files (*.rtf)|*.rtf|Text Files (*.txt)|*.txt"
        SaveFileDialog1.Title = "Save as..."
        SaveFileDialog1.FileName = "*.rtf"
        SaveFileDialog1.ShowDialog()

        Dim sFile As String = SaveFileDialog1.FileName

        rchBoxMain.SaveFile(sFile)

    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click

        SaveFileDialog1.Filter = "Richtext Files (*.rtf)|*.rtf|Text Files (*.txt)|*.txt"
        SaveFileDialog1.Title = "Save as..."
        SaveFileDialog1.FileName = "*.rtf"
        SaveFileDialog1.ShowDialog()

        Dim sFile As String = SaveFileDialog1.FileName

        rchBoxMain.SaveFile(sFile)

    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click

        OpenFileDialog1.Filter = "Richtext Files (*.rtf)|*.rtf|Text Files (*.txt)|*.txt"
        OpenFileDialog1.Title = "Select a Rich-Text Formatted File to Open..."
        OpenFileDialog1.FileName = "*.rtf"
        OpenFileDialog1.ShowDialog()

        Dim sFile As String = OpenFileDialog1.FileName

        If File.Exists(sFile) Then rchBoxMain.LoadFile(sFile)

    End Sub

    Private Sub ToolStripMenuItem14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem14.Click
        DoQuit()

        End
    End Sub

    Private Sub CutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CutToolStripMenuItem.Click
        rchBoxMain.Cut()
    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyToolStripMenuItem.Click
        rchBoxMain.Copy()
    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PasteToolStripMenuItem.Click
        rchBoxMain.Paste()
    End Sub

    Private Sub SelectAllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectAllToolStripMenuItem.Click
        rchBoxMain.SelectAll()
    End Sub


    Private Sub ToolStripMenuItem17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem17.Click
        rchBoxMain.Cut()
    End Sub

    Private Sub ToolStripMenuItem18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem18.Click
        rchBoxMain.Copy()
    End Sub

    Private Sub ToolStripMenuItem19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem19.Click
        rchBoxMain.Paste()
    End Sub

    Private Sub ToolStripMenuItem20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem20.Click
        rchBoxMain.SelectAll()
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged

        With rchBoxMain
            Dim oldpos As Integer = .SelectionStart

            If CheckBox2.CheckState = CheckState.Unchecked Then
                .SelectionStart = 1
                .SelectionLength = .TextLength
                .SelectionBackColor = .BackColor
                .SelectionLength = 0
                .SelectionStart = oldpos
            Else
                MatchLength()
                .SelectionLength = 0
                .SelectionStart = oldpos
            End If

        End With
    End Sub

    Private Sub ToolStripDropDownButton1_DropDownItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ToolStripDropDownButton1.DropDownItemClicked
        Dim index = 0
        Dim src = rchBoxMain.Text + Chr(32)
        Dim srcNew As String
        Dim maxIndex = src.Length
        Dim chandle() As Integer = GetChannelHandle()
        Dim cnames() As String = GetChannelNames()
        Dim cnt As Integer = Array.IndexOf(cnames, e.ClickedItem.Text)
        Dim results As New List(Of String)
        Dim cntSeg As Integer = 1
        Dim Segment As Integer

        If rchBoxMain.TextLength = 0 Then Exit Sub

        If rchBoxMain.TextLength >= 370 Then
            ToolStripProgressBar1.Increment(1)

            Do While index <= maxIndex

                Select Case cntSeg
                    Case 1 : Segment = 0
                    Case sndCount : Segment = 2
                    Case Else : Segment = 1
                End Select

                srcNew = src.Substring(index)
                Dim match = Regex.Match(srcNew, "^.{0," & 370 & "}\s").ToString()
                rchBoxMain.SelectionStart = index
                rchBoxMain.SelectionLength = match.Length

                If match.Length <= 1 Then
                    If cntSeg = sndCount Or match.Length = 0 Then
                        ToolStripProgressBar1.Value = 1
                        ToolStripProgressBar1.Visible = False
                        Exit Sub
                    End If

                    ToolStripProgressBar1.Increment(1)
                    results.Add(match)
                    index = index + match.Length
                    Continue Do
                End If

                RichTextBox1.Rtf = rchBoxMain.SelectedRtf
                ToolStripProgressBar1.Increment(1)
                mircSendText(ParseText(), chandle(cnt))
                cntSeg = cntSeg + 1
                Threading.Thread.Sleep(My.Settings.SendDelay * 100)
                results.Add(match)
                index = index + match.Length
            Loop
        Else
            RichTextBox1.Rtf = rchBoxMain.Rtf
            mircSendText(ParseText(), chandle(cnt))
        End If

        ToolStripProgressBar1.Visible = False
        ToolStripProgressBar1.Value = 1
        LastChan = cnames(cnt)
        If ToolStripMenuItem32.Checked = True Then Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub ToolStripDropDownButton1_DropDownOpening(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripDropDownButton1.DropDownOpening
        Dim cNames() As String = GetChannelNames()

        ToolStripDropDownButton1.DropDown.Items.Clear()

        For Each vb As String In cNames
            ToolStripDropDownButton1.DropDown.Items.Add(vb)
        Next
    End Sub

    Private Sub ToolStripMenuItem31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem31.Click
        If ToolStripMenuItem31.Checked = True Then
            Me.TopMost = True
        Else
            Me.TopMost = False
        End If

    End Sub
    Private Sub ToolStripMenuItem30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim mirc As Integer
        Dim mdiclient As Integer
        Dim Ret As Long
        Dim CurrentStyle As Long

        mirc = FindWindow("mirc", vbNullString)
        mdiclient = FindWindowEx(mirc, 0&, "mdiclient", vbNullString)
        SetParent(Me.Handle, mirc)
        CurrentStyle = GetWindowLong(Me.Handle, GWL_STYLE)
        Ret = SetWindowLong(Me.Handle, GWL_STYLE, CurrentStyle Or WS_POPUP)

        Me.Show()
    End Sub

    Private Sub ToolStripMenuItem21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem21.Click
        With My.Settings
            .colDefault = ComboBox1.ComboBox.SelectedIndex
            .colDialog = ComboBox3.SelectedIndex
            .colTelepathy = ComboBox4.SelectedIndex
            .colForeign = ComboBox5.SelectedIndex
            .colDone = ComboBox6.SelectedIndex
            .colNext = ComboBox7.SelectedIndex
            .colContinue = ComboBox8.SelectedIndex
            .colBegin = ComboBox9.SelectedIndex
            .txtDialog = TextBox1.Text
            .txtTelepathy = TextBox2.Text
            .txtForeign = TextBox3.Text
            .txtDone = TextBox4.Text
            .txtNext = TextBox5.Text
            .txtContinue = TextBox6.Text
            .txtBegin = TextBox7.Text
            .AlwaysTop = Me.TopMost
            .ClearSend = ToolStripMenuItem32.Checked
            .Save()
        End With
        Me.TopMost = False
        FMOptions.Show()
        FMOptions.TopMost = True
    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        With My.Settings
            .colDefault = ComboBox1.ComboBox.SelectedIndex
            .colDialog = ComboBox3.SelectedIndex
            .colTelepathy = ComboBox4.SelectedIndex
            .colForeign = ComboBox5.SelectedIndex
            .colDone = ComboBox6.SelectedIndex
            .colNext = ComboBox7.SelectedIndex
            .colContinue = ComboBox8.SelectedIndex
            .colBegin = ComboBox9.SelectedIndex
            .txtDialog = TextBox1.Text
            .txtTelepathy = TextBox2.Text
            .txtForeign = TextBox3.Text
            .txtDone = TextBox4.Text
            .txtNext = TextBox5.Text
            .txtContinue = TextBox6.Text
            .txtBegin = TextBox7.Text
            .AlwaysTop = Me.TopMost
            .ClearSend = ToolStripMenuItem32.Checked
            .Save()
        End With
        Me.TopMost = False
        FMOptions.Show()
        FMOptions.TopMost = True
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        rchBoxMain.SelectionColor = colTable(ComboBox1.SelectedIndex)
    End Sub


    Private Sub ToolStripMenuItem34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem34.Click
        Me.TopMost = False
        AboutBox1.Show()
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.CheckState = CheckState.Checked Then
            ComboBox1.Enabled = False
            If My.Settings.Standalone = False Then cmbProfiles.Enabled = True
        Else
            ComboBox1.Enabled = True
            If My.Settings.Standalone = False Then cmbProfiles.Enabled = False
        End If
    End Sub


    Private Sub RichTextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RichTextBox1.TextChanged

    End Sub

    Private Sub ComboBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.Click

    End Sub

    Private Sub cmbProfiles_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbProfiles.Click

    End Sub
End Class

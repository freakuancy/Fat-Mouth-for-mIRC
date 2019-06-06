Imports System.IO


Public Class FMOptions

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim searchPattern As String = "*.clr"

        With FolderBrowserDialog1
            .Description = "Voodoo 2 Script Directory..."
            .RootFolder = Environment.SpecialFolder.MyComputer
            .ShowDialog()

            Dim fPath As String
            fPath = .SelectedPath
            If Not Directory.Exists(fPath) Then
                TextBox1.Text = "None"
                Exit Sub
            End If
            Dim ndi As DirectoryInfo = New DirectoryInfo(fPath)
            Dim fnames() As FileInfo = ndi.GetFiles(searchPattern, SearchOption.AllDirectories)
            Try
                If File.Exists(fnames(0).FullName.ToString) Then
                    MsgBox("Found " + (UBound(fnames) + 1).ToString + " profiles under " + fPath.ToString + ".", vbOKOnly, "Profiles Found!")
                    TextBox1.Text = Path.GetDirectoryName(fnames(0).FullName.ToString)

                    GroupBox1.Enabled = False
                End If

            Catch ex As Exception
                MsgBox("No profiles were found within the selected directory or subdirectories. Until a valid Profile Path is specified, Fat Mouth will operate in Stand-Alone Mode.")
                TextBox1.Text = "None"
                GroupBox1.Enabled = True
            End Try
        End With
    End Sub

    Private Sub FMOptions_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed


    End Sub

    Private Sub FMOptions_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        e.Cancel = True
    End Sub

    Private Sub FMOptions_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        With My.Settings

            If My.Settings.FirstRun = False Then
                TextBox18.BackColor = .CanvasColor

                CheckBox1.Checked = .RememberChannel
                TrackBar1.Value = .SendDelay
                TextBox1.Text = .VPath
                TextBox20.Text = .CanvasFont.Name.ToString + ", " + .CanvasFont.Size.ToString
                TextBox2.BackColor = .Color0
                TextBox6.BackColor = .Color1
                TextBox10.BackColor = .Color2
                TextBox12.BackColor = .Color3
                TextBox3.BackColor = .Color4
                TextBox7.BackColor = .Color5
                TextBox11.BackColor = .Color6
                TextBox8.BackColor = .Color7
                TextBox4.BackColor = .Color8
                TextBox6.BackColor = .Color9
                TextBox9.BackColor = .Color10
                TextBox13.BackColor = .Color11
                TextBox17.BackColor = .Color12
                TextBox16.BackColor = .Color13
                TextBox15.BackColor = .Color14
                TextBox14.BackColor = .Color15
            Else
                TextBox20.Text = Me.Font.Name.ToString
            End If
        End With
    End Sub

    Private Sub TextBox3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.Click, TextBox3.Click, TextBox4.Click, TextBox5.Click, TextBox6.Click, TextBox7.Click, TextBox8.Click, TextBox9.Click, TextBox10.Click, TextBox11.Click, TextBox12.Click, TextBox13.Click, TextBox14.Click, TextBox15.Click, TextBox16.Click, TextBox17.Click, TextBox18.Click
        Dim txtBox As TextBox = sender
        With ColorDialog1
            .AllowFullOpen = True
            .AnyColor = True
            .SolidColorOnly = True
            .ShowDialog()
            Dim bckColor As Color = ColorDialog1.Color
            txtBox.BackColor = bckColor
        End With
    End Sub


    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        With FontDialog1
            .AllowVectorFonts = True
            .AllowSimulations = False
            .AllowScriptChange = False
            .FontMustExist = True
            Try
                .ShowDialog()
                Dim mFont As Font = .Font
                TextBox20.Text = mFont.Name.ToString + ", " + mFont.Size.ToString
                Label1.Font = mFont
                My.Settings.CanvasFont = mFont
            Catch EX As Exception
                MsgBox(EX.Message + vbNewLine + "Please try again.", vbOKOnly, "Error - Font Not Valid")
            End Try

        End With
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        With My.Settings

            .CanvasColor = TextBox18.BackColor
            .RememberChannel = CheckBox1.Checked
            .SendDelay = TrackBar1.Value
            .Color0 = TextBox2.BackColor
            .Color1 = TextBox6.BackColor
            .Color2 = TextBox10.BackColor
            .Color3 = TextBox12.BackColor
            .Color4 = TextBox3.BackColor
            .Color5 = TextBox7.BackColor
            .Color6 = TextBox11.BackColor
            .Color7 = TextBox8.BackColor
            .Color8 = TextBox4.BackColor
            .Color9 = TextBox6.BackColor
            .Color10 = TextBox9.BackColor
            .Color11 = TextBox13.BackColor
            .Color12 = TextBox17.BackColor
            .Color13 = TextBox16.BackColor
            .Color14 = TextBox15.BackColor
            .Color15 = TextBox14.BackColor
            .VPath = TextBox1.Text
            If TextBox1.Text = "None" Then
                .Standalone = True
            Else
                .Standalone = False
            End If


            .Save()
            .FirstRun = False
        End With
        Me.TopMost = False

        frmMain.TopMost = My.Settings.AlwaysTop
        FillForm()
        Me.Hide()
    End Sub

    Private Sub TextBox19_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox19.LostFocus
        If Integer.Parse(TextBox19.Text) > 400 Or Integer.Parse(TextBox19.Text) < 300 Then
            MsgBox("Please enter a number between 300 and 400.", vbOKOnly, "Chunk Length Invalid")
            TextBox19.Text = "370"
        End If
    End Sub

    Private Sub TextBox19_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox19.TextChanged

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TrackBar1_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar1.Scroll
        Label5.Text = "Send Delay: " & (TrackBar1.Value * 100) & "ms"

    End Sub
End Class
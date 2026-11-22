Public Class Form1
    Private Password As String
    Private FilePath As String

    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click
        If Not RichTextBox1.Text = "" Then
            If MsgBox("Do you want to clear all the text?", MsgBoxStyle.YesNo, "Editswfiles") = MsgBoxResult.Yes Then
                RichTextBox1.Text = ""
            End If
        End If
    End Sub

    Private Sub NewWindowToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewWindowToolStripMenuItem.Click
        Process.Start(Application.ExecutablePath)
    End Sub

    Private WasFileFound As Boolean = False

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        Dim DialogRe As DialogResult

        If WasFileFound = True Then
            DialogRe = DialogResult.OK
        Else
            DialogRe = OpenFileDialog1.ShowDialog()
        End If


        If DialogRe = DialogResult.OK Then

        ElseIf DialogRe = DialogResult.Cancel Then
            Return
        End If

        If OpenFileDialog1.FileName.EndsWith(".swfiles") Then
            'SWFiles (.swfiles) versions 1.0/1.1/2.0 and the new 3.0. 3.0 isn't build in the same way as past versions. (I imported)
            If IsSWFilse3(OpenFileDialog1.FileName) Then 'Checks if SWFiles verison is 3.0
                '3.0
                Dim writer As New SWFiles.CreateSWFiles
                Dim passform As New PasswordForm

                passform.ShowDialog()

                If passform.SkippedPassword = True Or passform.Password = "" Or Nothing Then
                    passform.Password = "Null"
                End If

                If My.Computer.FileSystem.DirectoryExists(My.Application.Info.DirectoryPath & "\TempFolder") Then
                    My.Computer.FileSystem.DeleteDirectory(My.Application.Info.DirectoryPath & "\TempFolder", FileIO.DeleteDirectoryOption.DeleteAllContents)
                    My.Computer.FileSystem.CreateDirectory(My.Application.Info.DirectoryPath & "\TempFolder")
                Else
                    My.Computer.FileSystem.CreateDirectory(My.Application.Info.DirectoryPath & "\TempFolder")
                End If

                Dim reader As String = writer.ReadSWFiles3File(OpenFileDialog1.FileName, My.Application.Info.DirectoryPath & "\TempFolder", passform.Password)

                If Not RichTextBox1.Text = "" Then
                    If MsgBox("Do you want to clear all the text and open the file?", MsgBoxStyle.YesNo, "Editswfiles") = MsgBoxResult.Yes Then
                        RichTextBox1.Text = reader
                    End If
                Else
                    RichTextBox1.Text = reader
                End If

                If My.Computer.FileSystem.DirectoryExists(My.Application.Info.DirectoryPath & "\TempFolder") Then
                    Try
                        My.Computer.FileSystem.DeleteDirectory(My.Application.Info.DirectoryPath & "\TempFolder", FileIO.DeleteDirectoryOption.DeleteAllContents)
                    Catch ex As Exception
                    End Try
                End If

                Me.Text = "Editswfiles - " & OpenFileDialog1.SafeFileName

            Else
                '1.0/1.1/2.0
                Dim TryCheck2 As Boolean = False
                Dim reader As String = My.Computer.FileSystem.ReadAllText(OpenFileDialog1.FileName)

                Try
                    Dim b As Byte() = Convert.FromBase64String(reader)
                    reader = System.Text.Encoding.UTF8.GetString(b)
                    TryCheck2 = True
                Catch ex As Exception
                End Try


                Try
                    Dim b As Byte() = Convert.FromBase64String(reader)
                    reader = System.Text.Encoding.UTF8.GetString(b)
                Catch ex As Exception
                End Try

                If Not RichTextBox1.Text = "" Then
                    If MsgBox("Do you want to clear all the text and open the file?", MsgBoxStyle.YesNo, "Editswfiles") = MsgBoxResult.Yes Then
                        RichTextBox1.Text = reader
                        Me.Text = "Editswfiles - " & OpenFileDialog1.SafeFileName
                    End If
                Else
                    RichTextBox1.Text = reader
                    Me.Text = "Editswfiles - " & OpenFileDialog1.SafeFileName
                End If


            End If




        ElseIf OpenFileDialog1.FileName.EndsWith(".txt") Then
            'Text File (.txt)
            If Not RichTextBox1.Text = "" Then
                If MsgBox("Do you want to clear all the text and open the file?", MsgBoxStyle.YesNo, "Editswfiles") = MsgBoxResult.Yes Then
                    RichTextBox1.Text = My.Computer.FileSystem.ReadAllText(OpenFileDialog1.FileName)
                    Me.Text = "Editswfiles - " & OpenFileDialog1.SafeFileName
                End If
            Else
                RichTextBox1.Text = My.Computer.FileSystem.ReadAllText(OpenFileDialog1.FileName)
                Me.Text = "Editswfiles - " & OpenFileDialog1.SafeFileName
            End If
        End If

    End Sub

    Private Function IsSWFilse3(filePath As String) As Boolean
        ' ZIP files start with "PK" (50 4B in hex)
        Dim buffer(3) As Byte
        Using fs As New IO.FileStream(filePath, IO.FileMode.Open, IO.FileAccess.Read)
            fs.Read(buffer, 0, buffer.Length)
        End Using

        ' Check first 4 bytes
        Return buffer(0) = &H50 AndAlso buffer(1) = &H4B AndAlso
               (buffer(2) = &H3 OrElse buffer(2) = &H5 OrElse buffer(2) = &H7) AndAlso
               (buffer(3) = &H4 OrElse buffer(3) = &H6 OrElse buffer(3) = &H8)
    End Function

    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
            If SaveFileDialog1.FilterIndex = 1 Then
                If My.Computer.FileSystem.FileExists(SaveFileDialog1.FileName) = False Then
                    My.Computer.FileSystem.WriteAllText(SaveFileDialog1.FileName, RichTextBox1.Text, False)
                End If
            ElseIf SaveFileDialog1.FilterIndex = 2 Then
                'swfiles 1.0
                If My.Computer.FileSystem.FileExists(SaveFileDialog1.FileName) = False Then
                    My.Computer.FileSystem.WriteAllText(SaveFileDialog1.FileName, RichTextBox1.Text, False)
                End If
            ElseIf SaveFileDialog1.FilterIndex = 3 Then
                'swfiles 1.1
                If My.Computer.FileSystem.FileExists(SaveFileDialog1.FileName) = False Then
                    Dim writer As String = RichTextBox1.Text
                    Dim byt As Byte() = System.Text.Encoding.UTF8.GetBytes(RichTextBox1.Text)
                    RichTextBox1.Text = Convert.ToBase64String(byt)
                    My.Computer.FileSystem.WriteAllText(SaveFileDialog1.FileName, writer, False)
                End If
            ElseIf SaveFileDialog1.FilterIndex = 4 Then
                'swfiles 2.0
                If My.Computer.FileSystem.FileExists(SaveFileDialog1.FileName) = False Then
                    Dim writer As String = RichTextBox1.Text
                    Dim byt As Byte() = System.Text.Encoding.UTF8.GetBytes(RichTextBox1.Text)
                    writer = Convert.ToBase64String(byt)
                    Dim byt2 As Byte() = System.Text.Encoding.UTF8.GetBytes(writer)
                    writer = Convert.ToBase64String(byt2)
                    My.Computer.FileSystem.WriteAllText(SaveFileDialog1.FileName, writer, False)
                End If
            ElseIf SaveFileDialog1.FilterIndex = 5 Then
                'swfiles 3.0
                Dim writer As New SWFiles.CreateSWFiles
                Dim passwordf As New PasswordForm
                passwordf.ShowDialog()

                If passwordf.SkippedPassword = True Or passwordf.Password = "" Or Nothing Then
                    passwordf.Password = "Null"
                End If

                If My.Computer.FileSystem.DirectoryExists(My.Application.Info.DirectoryPath & "\TempFolder") Then
                    My.Computer.FileSystem.DeleteDirectory(My.Application.Info.DirectoryPath & "\TempFolder", FileIO.DeleteDirectoryOption.DeleteAllContents)
                    My.Computer.FileSystem.CreateDirectory(My.Application.Info.DirectoryPath & "\TempFolder")
                Else
                    My.Computer.FileSystem.CreateDirectory(My.Application.Info.DirectoryPath & "\TempFolder")
                End If

                writer.CreateSWFiles3File(SaveFileDialog1.FileName, RichTextBox1.Text, My.Application.Info.DirectoryPath & "\TempFolder", passwordf.Password)

                If My.Computer.FileSystem.DirectoryExists(My.Application.Info.DirectoryPath & "\TempFolder") Then
                    Try
                        My.Computer.FileSystem.DeleteDirectory(My.Application.Info.DirectoryPath & "\TempFolder", FileIO.DeleteDirectoryOption.DeleteAllContents)
                    Catch ex As Exception
                    End Try
                End If
            ElseIf SaveFileDialog1.FilterIndex = 6 Then
                If My.Computer.FileSystem.FileExists(SaveFileDialog1.FileName) = False Then
                    My.Computer.FileSystem.WriteAllText(SaveFileDialog1.FileName, RichTextBox1.Text, False)
                End If
            End If
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        OpenFileDialog1.InitialDirectory = My.Application.Info.DirectoryPath

        RichTextBox1.AllowDrop = True
        AddHandler RichTextBox1.DragDrop, AddressOf DragAndDropEvent
    End Sub

    Private Sub DragAndDropEvent(sender As Object, e As DragEventArgs)
        Dim filename As Object = e.Data.GetData(DataFormats.Text)
        If filename IsNot Nothing Then
            RichTextBox1.AppendText(filename)
            Return
        End If


        Dim filename2 As Object = e.Data.GetData(DataFormats.FileDrop)

        If filename2 Is Nothing Then Return

        Dim list As String() = filename2

        Dim ev As New EventArgs


        Debug.WriteLine(list(0))
        WasFileFound = True
        OpenFileDialog1.FileName = list(0)
        OpenToolStripMenuItem_Click(OpenToolStripMenuItem, ev)

        WasFileFound = False
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Close()
    End Sub

    Private Sub UndoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UndoToolStripMenuItem.Click
        RichTextBox1.Undo()
    End Sub

    Private Sub RedoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RedoToolStripMenuItem.Click
        RichTextBox1.Redo()
    End Sub

    Private Sub CutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CutToolStripMenuItem.Click
        RichTextBox1.Cut()
    End Sub

    Private Sub CopyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem.Click
        RichTextBox1.Copy()
    End Sub

    Private Sub PasteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PasteToolStripMenuItem.Click
        RichTextBox1.Paste()
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click
        RichTextBox1.SelectedText = ""
    End Sub

    Private Sub SelectAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectAllToolStripMenuItem.Click
        RichTextBox1.SelectAll()
    End Sub

    Private Sub RichTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox1.KeyDown
        If RichTextBox1.Text = "" Then
            If e.KeyCode = Keys.Delete Then
                e.SuppressKeyPress = True
            ElseIf e.KeyCode = Keys.Back Then
                e.SuppressKeyPress = True
            End If
        End If
    End Sub

    Private Sub WordWrapToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WordWrapToolStripMenuItem.Click
        If RichTextBox1.WordWrap = True Then
            RichTextBox1.WordWrap = False
            WordWrapToolStripMenuItem.Checked = False
        ElseIf RichTextBox1.WordWrap = False Then
            RichTextBox1.WordWrap = True
            WordWrapToolStripMenuItem.Checked = True
        End If
    End Sub

    Private Sub FontToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FontToolStripMenuItem.Click
        If FontDialog1.ShowDialog() = DialogResult.OK Then
            RichTextBox1.Font = FontDialog1.Font
        End If
    End Sub

    Private Sub StatusBarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StatusBarToolStripMenuItem.Click
        If StatusStrip1.Visible = True Then
            StatusStrip1.Visible = False
            StatusBarToolStripMenuItem.Checked = False
        ElseIf StatusStrip1.Visible = False Then
            StatusStrip1.Visible = True
            StatusBarToolStripMenuItem.Checked = True
        End If
    End Sub

    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox1.TextChanged
        ToolStripStatusLabel1.Text = "Length: " & RichTextBox1.Text.Length.ToString()
    End Sub

    Private Sub AboutEditswfilesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutEditswfilesToolStripMenuItem.Click
        AbouEditswfiles.ShowDialog()
    End Sub
End Class

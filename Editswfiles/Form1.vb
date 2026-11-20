Public Class Form1
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

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
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

                    Dim reader As String = writer.ReadSWFiles3File(OpenFileDialog1.FileName, "Null", passform.Password)

                    If Not RichTextBox1.Text = "" Then
                        If MsgBox("Do you want to clear all the text and open the file?", MsgBoxStyle.YesNo, "Editswfiles") = MsgBoxResult.Yes Then
                            RichTextBox1.Text = reader
                        End If
                    Else
                        RichTextBox1.Text = reader
                    End If


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
                        End If
                    Else
                        RichTextBox1.Text = reader
                    End If
                End If




            ElseIf OpenFileDialog1.FileName.EndsWith(".txt") Then
                'Text File (.txt)
                If Not RichTextBox1.Text = "" Then
                    If MsgBox("Do you want to clear all the text and open the file?", MsgBoxStyle.YesNo, "Editswfiles") = MsgBoxResult.Yes Then
                        RichTextBox1.Text = My.Computer.FileSystem.ReadAllText(OpenFileDialog1.FileName)
                    End If
                End If

            Else

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

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        OpenFileDialog1.InitialDirectory = My.Application.Info.DirectoryPath
    End Sub
End Class

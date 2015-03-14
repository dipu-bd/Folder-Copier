Public Class Form1

    Dim succedd As Boolean = False

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            CheckedListBox1.Enabled = True
            CheckedListBox1.Items.Clear()

            For Each folder In FileIO.FileSystem.GetDirectories(TextBox2.Text)
                CheckedListBox1.Items.Add(folder)
            Next

            For Each file In FileIO.FileSystem.GetFiles(TextBox2.Text)
                CheckedListBox1.Items.Add(file)
            Next
            Me.Height = 510
            Me.FormBorderStyle = 4
        Else
            CheckedListBox1.Enabled = False
            Me.Height = 158
            Me.FormBorderStyle = 2
        End If
        Me.CenterToScreen()
    End Sub

    Private Sub TextBox1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TextBox1.MouseDoubleClick
        Process.Start(TextBox1.Text)
    End Sub

    Private Sub TextBox2_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TextBox2.MouseDoubleClick
        Process.Start(TextBox2.Text)
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        If My.Computer.FileSystem.DirectoryExists(TextBox2.Text) Then
            Button3.Enabled = True
            CheckBox1.Enabled = True
        Else
            Button3.Enabled = False
            CheckBox1.Enabled = False
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim folderBrowser1 As New FolderBrowserDialog
        If folderBrowser1.ShowDialog = 1 Then
            TextBox1.Text = folderBrowser1.SelectedPath
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim folderBrowser1 As New FolderBrowserDialog
        If folderBrowser1.ShowDialog = 1 Then
            TextBox2.Text = folderBrowser1.SelectedPath
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Button3.Enabled = False
        If Button3.Text = "Copy" Then
            If Not BackgroundWorker1.IsBusy Then
                BackgroundWorker1.RunWorkerAsync()
            End If
            Button3.Text = "Cancel"
            Button1.Enabled = False
            Button2.Enabled = False
            CheckBox1.Enabled = False
            CheckedListBox1.Enabled = False
            TextBox1.ReadOnly = True
            TextBox2.ReadOnly = True
        Else
            Button3.Text = "Copy"
            BackgroundWorker1.CancelAsync()
            Button1.Enabled = True
            Button2.Enabled = True
            CheckBox1.Enabled = True
            CheckedListBox1.Enabled = True
            TextBox1.ReadOnly = False
            TextBox2.ReadOnly = False
        End If
        Button3.Enabled = True
    End Sub

    Dim destination As String = ""
    Dim source As String = ""

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            If CheckBox1.Checked = True Then
                Dim checkedItems As CheckedListBox.CheckedItemCollection = CheckedListBox1.CheckedItems
                For Each item As String In checkedItems
                    source = item
                    Dim splt() As String = Split(item, "\")
                    destination = TextBox1.Text & "\" & splt(splt.Count - 1)
                    If My.Computer.FileSystem.FileExists(item) Then
                        My.Computer.FileSystem.CopyFile(source, destination, True)
                    ElseIf My.Computer.FileSystem.DirectoryExists(item) Then
                        My.Computer.FileSystem.CopyDirectory(source, destination, True)
                    End If
                Next
            Else
                My.Computer.FileSystem.CopyDirectory(TextBox2.Text, TextBox1.Text, True)
            End If
            succedd = True
        Catch
            succedd = False
        End Try
    End Sub

    Private Sub ToolStripStatusLabel1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ToolStripStatusLabel1.MouseMove
        Dim splt() As String = Split(source, "\")
        Try
            Dim srcLength As String
            Dim srcLen As Integer = My.Computer.FileSystem.GetFileInfo(source).Length / 1024
            If srcLen > 1024 Then
                srcLength = FormatNumber(srcLen / 1024, 2) & "MB"
            Else
                srcLength = FormatNumber(srcLen, 2) & "KB"
            End If

            ToolStripStatusLabel1.Text = splt(splt.Count - 1) & " : " & srcLength

        Catch ex As Exception
            ToolStripStatusLabel1.Text = splt(splt.Count - 1)
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Button3.Text = "Copy"
        Button1.Enabled = True
        Button2.Enabled = True
        CheckBox1.Enabled = True
        CheckedListBox1.Enabled = True
        TextBox1.ReadOnly = False
        TextBox2.ReadOnly = False
        If succedd = True Then
            Dim p As MsgBoxResult = MsgBox("Completed Successfully. Close now?", 4)
            Process.Start(TextBox1.Text)
            If p = 6 Then Me.Close()
        Else
            Dim p As MsgBoxResult = MsgBox("Copy Failed. Close now?", 4)
            Process.Start(TextBox1.Text)
            If p = 6 Then Me.Close()
        End If
    End Sub

    Private Sub CheckedListBox1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles CheckedListBox1.MouseDoubleClick
        Process.Start(CheckedListBox1.SelectedItem)
    End Sub

    Private Sub Form1_ResizeEnd(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeEnd
        If CheckBox1.Enabled = True Then Me.CenterToScreen()
    End Sub
End Class

#Region "ABOUT"
' / ------------------------------------------------------------------------------------------
' / Developer : Mr.Surapon Yodsanga (Thongkorn Tubtimkrob)
' / eMail : thongkorn@hotmail.com
' / URL: http://www.g2gnet.com (Khon Kaen - Thailand)
' / Facebook: http://www.facebook.com/g2gnet (for Thailand)
' / Facebook: http://www.facebook.com/commonindy (Worldwide)
' / More Info: http://www.g2gsoft.com
' /
' / Purpose: Lock/Unlock selected folder with code.
' / Microsoft Visual Basic .NET (2010)
' /
' / This is open source code under @CopyLeft by Thongkorn Tubtimkrob.
' / You can modify and/or distribute without to inform the developer.
' / ------------------------------------------------------------------------------------------
#End Region

Imports System.IO
Imports System.Security.AccessControl

Public Class frmLockUnlockFolder

    Dim strPath As String = Application.StartupPath.ToLower.Replace("bin\debug", "").Replace("bin\release", "") & "Images\"

    '// START HERE
    Private Sub frmLockUnlockFolder_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Label1.Text = ""
        With PictureBox1
            .SizeMode = PictureBoxSizeMode.StretchImage
            .Image = Image.FromFile(strPath & "people.png") '\\ Initialize to show sample image.
        End With
        '// Start Lock or set permission with code.
        Call LockFolder()
    End Sub

    '// Browse any folder.
    Private Sub btnFolderBrowse_Click(sender As System.Object, e As System.EventArgs) Handles btnFolderBrowse.Click
        Dim dlgFolderBrowse As New FolderBrowserDialog
        dlgFolderBrowse.SelectedPath = strPath
        '// If users select cancel then exit sub.
        If dlgFolderBrowse.ShowDialog() = Windows.Forms.DialogResult.Cancel Then Exit Sub
        '// 
        If Microsoft.VisualBasic.Right(dlgFolderBrowse.SelectedPath, 1) <> "\" Then dlgFolderBrowse.SelectedPath = dlgFolderBrowse.SelectedPath & "\"
        strPath = dlgFolderBrowse.SelectedPath
        Try
            '// Create ImageList dynamically.
            Dim img As New ImageList
            With img
                .ImageSize = New Point(128, 128)
                .ColorDepth = ColorDepth.Depth32Bit
            End With

            '// Before displaying the image, the folder must be Unlock first.
            Call UnlockFolder()

            Dim itemFolder As New List(Of ListViewItem)
            For Each imgFile In Directory.GetFiles(dlgFolderBrowse.SelectedPath.ToString())
                If imgFile.Contains(".png") Or imgFile.Contains(".jpg") Or imgFile.Contains(".bmp") Or imgFile.Contains(".gif") Then
                    img.Images.Add(Image.FromFile(imgFile))
                    itemFolder.Add(New ListViewItem(Path.GetFileName(imgFile)) With {.ImageIndex = img.Images.Count - 1})
                End If
            Next
            With ListView1
                .Items.Clear()
                .View = View.LargeIcon
                .LargeImageList = img
                .Items.AddRange(itemFolder.ToArray())
            End With

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Report Status", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
        '// When the image is displayed then Lock the folder again.
        Call LockFolder()
    End Sub

    Private Sub ListView1_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles ListView1.MouseClick
        Dim item As ListViewItem = ListView1.HitTest(e.Location).Item
        If item IsNot Nothing Then
            Me.PictureBox1.Image = Image.FromFile(strPath & item.Text)
            Label1.Text = "Image : " & strPath & item.Text
        End If
    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ListView1.SelectedIndexChanged
        'For Each item As ListViewItem In ListView1.SelectedItems
        'MessageBox.Show(item.Text)
        'Next
    End Sub

    ' / ------------------------------------------------------------------------------------------
    '// Set Permission for lock/unlock folder with code.
    Private Sub LockFolder()
        Dim fs As FileSystemSecurity = File.GetAccessControl(strPath)
        fs.AddAccessRule(New FileSystemAccessRule(Environment.UserName, FileSystemRights.FullControl, AccessControlType.Deny))
        File.SetAccessControl(strPath, fs)
    End Sub

    Private Sub UnlockFolder()
        Dim fs As FileSystemSecurity = File.GetAccessControl(strPath)
        fs.RemoveAccessRule(New FileSystemAccessRule(Environment.UserName, FileSystemRights.FullControl, AccessControlType.Deny))
        File.SetAccessControl(strPath, fs)
    End Sub
    ' / ------------------------------------------------------------------------------------------

    Private Sub frmLockUnlockFolder_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Call UnlockFolder()    '// Unlock before the end of the program.
        Me.Dispose()
        GC.SuppressFinalize(Me)
        Application.Exit()
    End Sub

    Private Sub btnExit_Click(sender As System.Object, e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub
End Class

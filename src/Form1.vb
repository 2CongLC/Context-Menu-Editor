Imports System
Imports System.Text
Imports System.IO
Imports Microsoft.Win32

Public Class Form1
    Private Reg As RegistryKey

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = 1
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            TextBox1.Text = OpenFileDialog1.FileName.Trim
            TextBox3.Text = Path.GetFileName(TextBox1.Text)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If OpenFileDialog2.ShowDialog = DialogResult.OK Then
            TextBox2.Text = OpenFileDialog2.FileName.Trim
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            If File.Exists(TextBox1.Text) = True OrElse File.Exists(TextBox2.Text) = True Then
                If String.IsNullOrWhiteSpace(TextBox3.Text) = False Then

                    Registry.ClassesRoot.OpenSubKey("Directory\Background\shell\" + TextBox3.Text, True)
                    If Reg Is Nothing Then My.Computer.Registry.ClassesRoot.CreateSubKey("Directory\Background\shell\" + TextBox3.Text, True)
                    My.Computer.Registry.ClassesRoot.OpenSubKey("Directory\Background\shell\" + TextBox3.Text, True).SetValue("Icon", Chr(34) & TextBox1.Text & Chr(34), RegistryValueKind.String)
                    Dim p As String = String.Empty
                    If ComboBox1.SelectedIndex = 0 Then
                        p = "top"
                    ElseIf ComboBox1.SelectedIndex = 1 Then
                        p = "midle"
                    ElseIf ComboBox1.SelectedIndex = 2 Then
                        p = "bottom"
                    End If
                    My.Computer.Registry.ClassesRoot.OpenSubKey("Directory\Background\shell\" + TextBox3.Text, True).SetValue("Position", p, RegistryValueKind.String)
                    My.Computer.Registry.ClassesRoot.CreateSubKey("Directory\Background\shell\" + TextBox3.Text + "\" + "Command\").SetValue("", Chr(34) & TextBox2.Text & Chr(34), RegistryValueKind.String)

                    MessageBox.Show("Đã thêm vào menu !!!")
                Else
                    MessageBox.Show("Nhập tên ứng dụng trong menu !")
                End If
            Else
                MessageBox.Show("Hãy chọn đúng đường dẫn ứng dụng và biểu tượng !")
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            My.Computer.Registry.ClassesRoot.DeleteSubKeyTree("Directory\Background\shell\" + TextBox3.Text, True)
            MessageBox.Show("Đã xóa khỏi menu !!!")
        Catch ex As Exception

        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            Process.Start("https://2conglc-vn.blogspot.com/2020/08/cong-cu-chinh-sua-menu-chuot-phai.html")
        Catch ex As Exception

        End Try
    End Sub
End Class

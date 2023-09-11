Public Class Form2
    Dim prote As New My.MySettings()

    Private Sub login()
        If TextBox1.Text = "Srecko" And TextBox2.Text = prote.pass Then
            Me.Hide()
            MsgBox("You are now logged as an admin", MsgBoxStyle.OkOnly)
            Form1.Show()
        ElseIf TextBox1.Text = "User" And TextBox2.Text = prote.passuser Then
            Me.Hide()
            MsgBox("You are now logged as an user", MsgBoxStyle.OkOnly)
            Form1.Show()
            Form1.updatebtn.Enabled = False
            Form1.retrievebtn.Enabled = False
            Form1.deletebtn.Enabled = False
            Form1.addbtn.Enabled = False
            Form1.fororder.Enabled = False
        Else
            MsgBox("Incorrect username or password", MsgBoxStyle.OkOnly)
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox1.Focus()

        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        login()
    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        If Asc(e.KeyChar) = 13 Then

            login()

        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked

        TextBox1.Text = ""
        TextBox2.Text = ""
        Me.Hide()
        MsgBox("Password Changer", MsgBoxStyle.OkOnly)
        Form3.Show()

    End Sub
End Class
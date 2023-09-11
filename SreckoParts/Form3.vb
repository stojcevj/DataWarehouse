Public Class Form3
    Dim protection As New My.MySettings()

    Private Sub Change()

        If TextBox1.Text = "User" And TextBox2.Text = protection.seccode Then

            protection.passuser = TextBox3.Text
            protection.Save()
            MsgBox("Your password was changed", MsgBoxStyle.OkOnly)
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            Me.Hide()
            Form2.TextBox1.Text = ""
            Form2.TextBox2.Text = ""
            Form2.Show()

        ElseIf TextBox1.Text = "Srecko" And TextBox2.Text = protection.seccode Then

            protection.pass = TextBox3.Text
            protection.Save()
            MsgBox("Your password was changed", MsgBoxStyle.OkOnly)
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            Me.Hide()
            Form2.TextBox1.Text = ""
            Form2.TextBox2.Text = ""
            Form2.Show()

        Else

            MsgBox("Inncorect user or security code", MsgBoxStyle.Critical)
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox1.Focus()


        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Change()


    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Hide()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        Form2.TextBox1.Text = ""
        Form2.TextBox2.Text = ""
        Form2.Show()
    End Sub

    Private Sub TextBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        If Asc(e.KeyChar) = 13 Then

            Change()


        End If
    End Sub

    Private Sub Form3_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Form2.Close()
    End Sub
End Class
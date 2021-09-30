Imports System.Net.Mail
Public Class Form6
    Dim mesaj As New MailMessage()
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim cvp As New DialogResult
        cvp = MessageBox.Show("Mail göndermek istediğinizden emin misiniz?", "Emin misiniz?", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If cvp = DialogResult.Yes Then
            Try
                mesaj.From = New MailAddress(TextBox1.Text)
                mesaj.To.Add(TextBox3.Text)
                mesaj.Subject = TextBox5.Text
                mesaj.Body = TextBox4.Text
                If RadioButton1.Checked = True And RadioButton2.Checked = False Then
                    Dim Smtp As New SmtpClient("smtp.gmail.com")
                    Smtp.EnableSsl = True
                    Smtp.Port = 587
                    Smtp.Credentials = New Net.NetworkCredential(TextBox1.Text, TextBox2.Text)
                    Smtp.Send(mesaj)
                ElseIf RadioButton2.Checked = True And RadioButton1.Checked = False Then
                    Dim Smtp As New SmtpClient("smtp.live.com")
                    Smtp.EnableSsl = True
                    Smtp.Port = 587
                    Smtp.Credentials = New Net.NetworkCredential(TextBox1.Text, TextBox2.Text)
                    Smtp.Send(mesaj)
                End If

                MessageBox.Show("Mesajınız iletildi!", "Mail", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show("Mesajınız iletilemedi verileri kontrol ediniz!", "Mail", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        OpenFileDialog1.InitialDirectory = "C:\Users\kacan\Desktop"
        OpenFileDialog1.FileName = "maliyet.xlsx"
        OpenFileDialog1.ShowDialog()
        mesaj.Attachments.Add(New Attachment(OpenFileDialog1.FileName))

    End Sub
    Private Sub Form6_Leave(sender As Object, e As EventArgs) Handles MyBase.Leave
        Me.Close()
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim frm1 As New Form1
        Me.Hide()
        frm1.ShowDialog()
        Me.Close()
    End Sub
End Class
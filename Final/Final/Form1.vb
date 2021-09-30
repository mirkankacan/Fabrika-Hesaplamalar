Imports hesap
Imports System.Data.SqlClient
Public Class Form1
    Dim conn As New SqlConnection("server=xxx;user=sa;pwd=123456;database=fabrika;")
    Dim cmd As New SqlCommand
    Dim adap As New SqlDataAdapter
    Dim dt As New DataTable
    Dim bs As New BindingSource

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim dll As New hesap.Class1
        Dim cvp As New DialogResult
        If TextBox1.Text <> "" And TextBox2.Text <> "" And TextBox3.Text <> "" And TextBox4.Text <> "" And TextBox5.Text <> "" And TextBox6.Text <> "" And TextBox7.Text <> "" And TextBox8.Text <> "" And TextBox9.Text <> "" And TextBox10.Text <> "" And TextBox11.Text <> "" And TextBox12.Text <> "" And TextBox13.Text <> "" And TextBox14.Text <> "" And TextBox15.Text <> "" Then
            Dim sonuc As Double = dll.hesapla(Convert.ToDouble(TextBox1.Text), Convert.ToDouble(TextBox2.Text), Convert.ToDouble(TextBox3.Text), Convert.ToDouble(TextBox4.Text), Convert.ToDouble(TextBox5.Text), Convert.ToDouble(TextBox6.Text), Convert.ToDouble(TextBox7.Text), Convert.ToDouble(TextBox8.Text), Convert.ToDouble(TextBox9.Text), Convert.ToDouble(TextBox10.Text), Convert.ToDouble(TextBox11.Text), Convert.ToDouble(TextBox12.Text), Convert.ToDouble(TextBox13.Text), Convert.ToDouble(TextBox14.Text), Convert.ToDouble(TextBox15.Text))
            cvp = MessageBox.Show("Hesaplama yapmaktan ve veritabanına kaydetmekten emin misiniz?", "Emin misiniz?", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If cvp = DialogResult.Yes Then
                conn.Open()
                cmd.Connection = conn
                cmd.CommandType = CommandType.Text
                cmd.CommandText = "insert into harcama (Toplam_Malzeme,Hammade_Satın_Alma_Maliyeti,Finansman_Maliyeti,Fire_Miktarı,Fire_Oranı,Makine_Aylık_Amortismanı,Makine_Çalışma_Süresi,Makinenin_Harcadığı_Enerji_Maliyeti,Ürün_İşlenme_Süresi,Brüt_Ücret,SSK_Maliyeti,Yemek_ve_Taşıma_Ücreti,İş_Güvenlik_Harcamaları,Ödenen_Toplam_Maaş,Üretim_Adeti,Toplam_Maliyet) values(@a,@b,@c,@d,@e,@f,@g,@h,@o,@j,@k,@l,@m,@n,@z,@x)"
                cmd.Parameters.Add("@a", SqlDbType.Float).Value = Convert.ToDouble(TextBox1.Text)
                cmd.Parameters.Add("@b", SqlDbType.Float).Value = Convert.ToDouble(TextBox2.Text)
                cmd.Parameters.Add("@c", SqlDbType.Float).Value = Convert.ToDouble(TextBox3.Text)
                cmd.Parameters.Add("@d", SqlDbType.Float).Value = Convert.ToDouble(TextBox4.Text)
                cmd.Parameters.Add("@e", SqlDbType.Float).Value = Convert.ToDouble(TextBox5.Text)
                cmd.Parameters.Add("@f", SqlDbType.Float).Value = Convert.ToDouble(TextBox6.Text)
                cmd.Parameters.Add("@g", SqlDbType.Float).Value = Convert.ToDouble(TextBox7.Text)
                cmd.Parameters.Add("@h", SqlDbType.Float).Value = Convert.ToDouble(TextBox8.Text)
                cmd.Parameters.Add("@o", SqlDbType.Float).Value = Convert.ToDouble(TextBox9.Text)
                cmd.Parameters.Add("@j", SqlDbType.Float).Value = Convert.ToDouble(TextBox10.Text)
                cmd.Parameters.Add("@k", SqlDbType.Float).Value = Convert.ToDouble(TextBox11.Text)
                cmd.Parameters.Add("@l", SqlDbType.Float).Value = Convert.ToDouble(TextBox12.Text)
                cmd.Parameters.Add("@m", SqlDbType.Float).Value = Convert.ToDouble(TextBox13.Text)
                cmd.Parameters.Add("@n", SqlDbType.Float).Value = Convert.ToDouble(TextBox14.Text)
                cmd.Parameters.Add("@z", SqlDbType.Float).Value = Convert.ToDouble(TextBox15.Text)
                cmd.Parameters.Add("@x", SqlDbType.Float).Value = Convert.ToDouble(sonuc)
                cmd.ExecuteNonQuery()
                conn.Close()
                MessageBox.Show("Toplam maliyet " & sonuc & " TL", "Maliyet", MessageBoxButtons.OK, MessageBoxIcon.Information)
                MessageBox.Show("Hesaplama başarılı bir şekilde yapıldı ve veritabanına kaydedildi!", "Bilgilendirme", MessageBoxButtons.OK, MessageBoxIcon.Information)
                cmd.Parameters.Clear()
                TextBox1.Clear()
                TextBox2.Clear()
                TextBox3.Clear()
                TextBox4.Clear()
                TextBox5.Clear()
                TextBox6.Clear()
                TextBox7.Clear()
                TextBox8.Clear()
                TextBox9.Clear()
                TextBox10.Clear()
                TextBox11.Clear()
                TextBox12.Clear()
                TextBox13.Clear()
                TextBox14.Clear()
                TextBox15.Clear()
            End If
        Else
            MessageBox.Show("Alanları doğru bir şekilde doldurunuz!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

    End Sub
    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox9.KeyPress, TextBox8.KeyPress, TextBox7.KeyPress, TextBox6.KeyPress, TextBox5.KeyPress, TextBox4.KeyPress, TextBox3.KeyPress, TextBox2.KeyPress, TextBox15.KeyPress, TextBox14.KeyPress, TextBox13.KeyPress, TextBox12.KeyPress, TextBox11.KeyPress, TextBox10.KeyPress, TextBox1.KeyPress

        If e.KeyChar <> ControlChars.Back Then
            e.Handled = Not (Char.IsDigit(e.KeyChar) Or e.KeyChar = ".")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim frm2 As New Form2
        Me.Hide()
        frm2.ShowDialog()
        Me.Close()
    End Sub

    Private Sub Form1_Leave(sender As Object, e As EventArgs) Handles MyBase.Leave
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim frm3 As New Form3
        Me.Hide()
        frm3.ShowDialog()
        Me.Close()
    End Sub


    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim frm5 As New Form5
        Me.Hide()
        frm5.ShowDialog()
        Me.Close()
    End Sub

    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click
        Dim frm4 As New Form4
        Me.Hide()
        frm4.ShowDialog()
        Me.Close()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim frm6 As New Form6
        Me.Hide()
        frm6.ShowDialog()
        Me.Close()

    End Sub
End Class

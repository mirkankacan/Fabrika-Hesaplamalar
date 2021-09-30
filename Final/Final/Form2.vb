Imports System.Data.SqlClient
Public Class Form2
    Dim conn As New SqlConnection("server=xxx;user=sa;pwd=123456;database=fabrika;")
    Dim cmd As New SqlCommand
    Dim adap As New SqlDataAdapter
    Dim dt As New DataTable
    Dim bs As New BindingSource

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.AllowUserToResizeRows = False
        DataGridView1.AllowUserToResizeColumns = False
        TextBox16.Visible = False
        TextBox17.Enabled = False
        conn.Open()
        cmd.Connection = conn
        cmd.CommandText = "select * from harcama"
        cmd.CommandType = CommandType.Text
        cmd.ExecuteScalar()
        adap.SelectCommand = cmd
        adap.Fill(dt)
        bs.DataSource = dt
        DataGridView1.DataSource = bs
        TextBox1.DataBindings.Add("Text", bs, "Toplam_Malzeme")
        TextBox2.DataBindings.Add("Text", bs, "Hammade_Satın_Alma_Maliyeti")
        TextBox3.DataBindings.Add("Text", bs, "Finansman_Maliyeti")
        TextBox4.DataBindings.Add("Text", bs, "Fire_Miktarı")
        TextBox5.DataBindings.Add("Text", bs, "Fire_Oranı")
        TextBox6.DataBindings.Add("Text", bs, "Makine_Aylık_Amortismanı")
        TextBox7.DataBindings.Add("Text", bs, "Makine_Çalışma_Süresi")
        TextBox8.DataBindings.Add("Text", bs, "Makinenin_Harcadığı_Enerji_Maliyeti")
        TextBox9.DataBindings.Add("Text", bs, "Ürün_İşlenme_Süresi")
        TextBox10.DataBindings.Add("Text", bs, "Brüt_Ücret")
        TextBox11.DataBindings.Add("Text", bs, "SSK_Maliyeti")
        TextBox12.DataBindings.Add("Text", bs, "Yemek_ve_Taşıma_Ücreti")
        TextBox13.DataBindings.Add("Text", bs, "İş_Güvenlik_Harcamaları")
        TextBox14.DataBindings.Add("Text", bs, "Ödenen_Toplam_Maaş")
        TextBox15.DataBindings.Add("Text", bs, "Üretim_Adeti")
        TextBox16.DataBindings.Add("Text", bs, "ID")
        TextBox17.DataBindings.Add("Text", bs, "Toplam_Maliyet")
        conn.Close()
    End Sub
    Sub refresh()
        conn.Open()
        cmd.Connection = conn
        cmd.CommandText = "select * from harcama"
        cmd.CommandType = CommandType.Text
        cmd.ExecuteScalar()
        adap.SelectCommand = cmd
        adap.Fill(dt)
        DataGridView1.DataSource = dt
        conn.Close()
    End Sub

    Private Sub Form2_Leave(sender As Object, e As EventArgs) Handles MyBase.Leave
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text <> "" And TextBox2.Text <> "" And TextBox3.Text <> "" And TextBox4.Text <> "" And TextBox5.Text <> "" And TextBox6.Text <> "" And TextBox7.Text <> "" And TextBox8.Text <> "" And TextBox9.Text <> "" And TextBox10.Text <> "" And TextBox11.Text <> "" And TextBox12.Text <> "" And TextBox13.Text <> "" And TextBox14.Text <> "" And TextBox15.Text <> "" Then
            conn.Open()
            cmd.Connection = conn
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "update harcama set Toplam_Malzeme=@a,Hammade_Satın_Alma_Maliyeti=@b,Finansman_Maliyeti=@c,Fire_Miktarı=@d,Fire_Oranı=@e,Makine_Aylık_Amortismanı=@f,Makine_Çalışma_Süresi=@g,Makinenin_Harcadığı_Enerji_Maliyeti=@h,Ürün_İşlenme_Süresi=@o,Brüt_Ücret=@j,SSK_Maliyeti=@k,Yemek_ve_Taşıma_Ücreti=@l,İş_Güvenlik_Harcamaları=@m,Ödenen_Toplam_Maaş=@n,Üretim_Adeti=@z where ID=@id"
            cmd.Parameters.Add("@id", SqlDbType.Int).Value = Convert.ToInt32(TextBox16.Text)
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
            cmd.ExecuteNonQuery()
            conn.Close()
            MessageBox.Show("Kayıt başarıyla güncellendi!", "Güncelleme", MessageBoxButtons.OK, MessageBoxIcon.Information)
            DataGridView1.DataSource = Nothing
            dt.Rows.Clear()
            dt.Clear()
            DataGridView1.Rows.Clear()
            DataGridView1.Refresh()
            DataGridView1.Refresh()
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
            TextBox16.Clear()
            TextBox17.Clear()
            refresh()


        Else
            MessageBox.Show("Alanları doğru girdiğinizden emin olunuz!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

    End Sub
    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox9.KeyPress, TextBox8.KeyPress, TextBox7.KeyPress, TextBox6.KeyPress, TextBox5.KeyPress, TextBox4.KeyPress, TextBox3.KeyPress, TextBox2.KeyPress, TextBox15.KeyPress, TextBox14.KeyPress, TextBox13.KeyPress, TextBox12.KeyPress, TextBox11.KeyPress, TextBox10.KeyPress, TextBox1.KeyPress
        If e.KeyChar <> ControlChars.Back Then
            e.Handled = Not (Char.IsDigit(e.KeyChar) Or e.KeyChar = ".")
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim frm1 As New Form1
        Me.Hide()
        frm1.ShowDialog()
        Me.Close()
    End Sub

End Class
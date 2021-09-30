Imports System.Data.SqlClient
Public Class Form3
    Dim conn As New SqlConnection("server=xxx;user=sa;pwd=123456;database=fabrika;")
    Dim cmd As New SqlCommand
    Dim adap As New SqlDataAdapter
    Dim dt As New DataTable
    Dim dr As SqlDataReader
    Dim bs As New BindingSource
    Dim durum As New Boolean
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.AllowUserToResizeRows = False
        DataGridView1.AllowUserToResizeColumns = False
        conn.Open()
        cmd.Connection = conn
        cmd.CommandText = "select * from harcama"
        cmd.CommandType = CommandType.Text
        cmd.ExecuteScalar()
        adap.SelectCommand = cmd
        adap.Fill(dt)
        bs.DataSource = dt
        DataGridView1.DataSource = bs
        TextBox1.DataBindings.Add("Text", bs, "ID")
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

    Sub a()
        If conn.State = ConnectionState.Closed Then
            conn.Open()
            cmd.Connection = conn
            cmd.CommandText = "select * from harcama where ID='" + TextBox1.Text + "'"
            cmd.CommandType = CommandType.Text
            dr = cmd.ExecuteReader()
            If dr.Read() Then
                durum = False
            Else
                durum = True
            End If
            conn.Close()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        a()
        If durum = False Then
            Dim cvp As New DialogResult
            cvp = MessageBox.Show("Veriyi veritabanından silmekten emin misiniz?", "Emin misiniz?", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If cvp = DialogResult.Yes Then
                conn.Open()
                cmd.Connection = conn
                cmd.CommandType = CommandType.Text
                cmd.CommandText = "delete from harcama where ID='" + TextBox1.Text + "'"
                cmd.ExecuteNonQuery()
                conn.Close()
                MessageBox.Show("Kayıt başarıyla silindi!", "Silme", MessageBoxButtons.OK, MessageBoxIcon.Information)
                DataGridView1.DataSource = Nothing
                dt.Rows.Clear()
                dt.Clear()
                DataGridView1.Rows.Clear()
                DataGridView1.Refresh()
                DataGridView1.Refresh()
                refresh()
            End If
        Else
            MessageBox.Show("Bu ID'de herhangi bir kayıt bulunmuyor!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

    End Sub
    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar <> ControlChars.Back Then
            e.Handled = Not (Char.IsDigit(e.KeyChar))
        End If
    End Sub

    Private Sub Form3_Leave(sender As Object, e As EventArgs) Handles MyBase.Leave
        Me.Close()

    End Sub

    Private Sub LinkLabel1_LinkClicked_1(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim frm1 As New Form1
        Me.Hide()
        frm1.ShowDialog()
        Me.Close()
    End Sub
End Class
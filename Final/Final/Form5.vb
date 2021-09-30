Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel

Public Class Form5
    Dim conn As New SqlConnection("server=xxx;user=sa;pwd=123456;database=fabrika;")
    Dim cmd As New SqlCommand
    Dim adap As New SqlDataAdapter
    Dim dt As New DataTable
    Dim exc As Excel.Application
    Dim kitap As Excel.Workbook
    Dim sayfa As Excel.Worksheet

    Private Sub Form5_Load(sender As Object, e As EventArgs) Handles MyBase.Load

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

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim frm1 As New Form1
        Me.Hide()
        frm1.ShowDialog()
        Me.Close()

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        exc = CreateObject("Excel.Application")
        exc.Visible = True
        kitap = exc.Workbooks.Add
        sayfa = exc.ActiveWorkbook.ActiveSheet
        sayfa.Cells(1, 1) = "ID"
        sayfa.Cells(1, 2) = "TOPLAM MALZEME"
        sayfa.Cells(1, 3) = "HAMMADE SATIN ALINMA MALİYETİ"
        sayfa.Cells(1, 4) = "FİNANSMAN MALİYETİ"
        sayfa.Cells(1, 5) = "FİRE MİKTARI"
        sayfa.Cells(1, 6) = "FİRE ORANI"
        sayfa.Cells(1, 7) = "MAKİNE AYLIK AMORTİSMANI"
        sayfa.Cells(1, 8) = "MAKİNE ÇALIŞMA SÜRESİ"
        sayfa.Cells(1, 9) = "MAKİNENİN HARCADIĞI ENERJİ MALİYETİ"
        sayfa.Cells(1, 10) = "ÜRÜNÜN İŞLENME SÜRESİ"
        sayfa.Cells(1, 11) = "BRÜT ÜCRET"
        sayfa.Cells(1, 12) = "SSK MALİYETİ"
        sayfa.Cells(1, 13) = "YEMEK VE TAŞIMA ÜCRETİ"
        sayfa.Cells(1, 14) = "İŞ GÜVENLİK HARCAMALARI"
        sayfa.Cells(1, 15) = "ÖDENEN TOPLAM MAAŞ"
        sayfa.Cells(1, 16) = "ÜRETİM ADETİ"
        sayfa.Cells(1, 17) = "TOPLAM MALİYET"
        Dim i As Integer
        For i = 1 To DataGridView1.RowCount
            sayfa.Cells(i + 1, 1).Value = DataGridView1.Rows(i - 1).Cells(0).Value
            sayfa.Cells(i + 1, 2).Value = DataGridView1.Rows(i - 1).Cells(1).Value
            sayfa.Cells(i + 1, 3).Value = DataGridView1.Rows(i - 1).Cells(2).Value
            sayfa.Cells(i + 1, 4).Value = DataGridView1.Rows(i - 1).Cells(3).Value
            sayfa.Cells(i + 1, 5).Value = DataGridView1.Rows(i - 1).Cells(4).Value
            sayfa.Cells(i + 1, 6).Value = DataGridView1.Rows(i - 1).Cells(5).Value
            sayfa.Cells(i + 1, 7).Value = DataGridView1.Rows(i - 1).Cells(6).Value
            sayfa.Cells(i + 1, 8).Value = DataGridView1.Rows(i - 1).Cells(7).Value
            sayfa.Cells(i + 1, 9).Value = DataGridView1.Rows(i - 1).Cells(8).Value
            sayfa.Cells(i + 1, 10).Value = DataGridView1.Rows(i - 1).Cells(9).Value
            sayfa.Cells(i + 1, 11).Value = DataGridView1.Rows(i - 1).Cells(10).Value
            sayfa.Cells(i + 1, 12).Value = DataGridView1.Rows(i - 1).Cells(11).Value
            sayfa.Cells(i + 1, 13).Value = DataGridView1.Rows(i - 1).Cells(12).Value
            sayfa.Cells(i + 1, 14).Value = DataGridView1.Rows(i - 1).Cells(13).Value
            sayfa.Cells(i + 1, 15).Value = DataGridView1.Rows(i - 1).Cells(14).Value
            sayfa.Cells(i + 1, 16).Value = DataGridView1.Rows(i - 1).Cells(15).Value
            sayfa.Cells(i + 1, 17).Value = DataGridView1.Rows(i - 1).Cells(16).Value
        Next

    End Sub
    Private Sub Form5_Leave(sender As Object, e As EventArgs) Handles MyBase.Leave
        Me.Close()
    End Sub

End Class
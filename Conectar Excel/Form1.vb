Public Class Form1
    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        Dim Open As System.IO.Stream
        Open = OpenFileDialog1.OpenFile
        TextBox1.Text = OpenFileDialog1.FileName
        If Not (Open Is Nothing) Then
            Open.Close()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "Seleccionar archivo Excel.xlsx"
        OpenFileDialog1.ShowDialog()
    End Sub

    Sub CARGAREXCEL()
        Try
            Dim Conn As System.Data.OleDb.OleDbConnection
            Dim ds As System.Data.DataSet
            Dim command As System.Data.OleDb.OleDbDataAdapter

            Conn = New System.Data.OleDb.OleDbConnection(String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='{0}' ;Extended Properties='Excel 12.0;HDR=YES;'", TextBox1.Text))
            command = New System.Data.OleDb.OleDbDataAdapter("SELECT * FROM [Hoja1$]", Conn)
            command.TableMappings.Add("Table", "Net-informations.com")
            ds = New System.Data.DataSet
            command.Fill(ds)
            DataGridView1.DataSource = ds.Tables(0)
            Conn.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        CARGAREXCEL()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Form2.Show()
        Button1.Enabled = False
        Button2.Text = "ACTUALIZAR"
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form3.Show()
        Button1.Enabled = False
        Button2.Text = "ACTUALIZAR"
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Form4.Show()
        Button1.Enabled = False
        Button2.Text = "ACTUALIZAR"
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = "" Then
            Button3.Enabled = False
            Button4.Enabled = False
            Button5.Enabled = False
            Button1.Enabled = True
            Button2.Text = "CARGAR"
            Form2.Hide()
            Form3.Hide()
            Form4.Hide()
            DataGridView1.DataSource = vbEmpty
        Else
            Button3.Enabled = True
            Button4.Enabled = True
            Button5.Enabled = True
        End If
    End Sub
End Class

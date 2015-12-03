Imports Microsoft.Office.Interop

Public Class Form4
    Dim conversion As String
    Dim cadena As String
    Dim dato As String
    Dim dato2 As String

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For i As Integer = 1 To 1  '48576
            For j = 65 To 90
                'For x As Integer = 65 To 90
                conversion = Chr(j) & i
                cadena = conversion
                ComboBox1.Items.Add(cadena)
                'Next x
            Next j
        Next i
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        dato = TextBox1.Text
        dato2 = ComboBox1.Text

        If (dato <> "") And (dato2 <> "") Then
            Dim crearExcel As Excel.Application = CType(CreateObject("Excel.Application"), Excel.Application)
            crearExcel.Visible = True
            Dim ubExcel As Excel.Workbook = crearExcel.Workbooks.Open(Form1.TextBox1.Text)
            Dim hojaExcel As Excel.Worksheet = DirectCast(ubExcel.Worksheets("Hoja1"), Excel.Worksheet)

            hojaExcel.Range(ComboBox1.Text).Value = TextBox1.Text

            ubExcel.Close(SaveChanges:=True)
            crearExcel.Quit()

            'MsgBox("COLUMNA CREADA EXITOSAMENTE, ACTUALICE VENTANA PRINCIPAL", MsgBoxStyle.Information, "CONFIRMACIÓN")
            Form1.Show()

            TextBox1.Clear()
            ComboBox1.Text = ""
            TextBox1.Focus()
        Else
            MsgBox("INGRESE DATO", MsgBoxStyle.Critical, "ERROR")
            TextBox1.Clear()
            TextBox1.Focus()
        End If
    End Sub
End Class
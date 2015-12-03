Imports Microsoft.Office.Interop

Public Class Form3
    Dim conversion As String
    Dim cadena As String
    Dim dato As String

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        dato = ComboBox1.Text

        If ComboBox1.Text <> "" Then
            Dim crearExcel As Excel.Application = CType(CreateObject("Excel.Application"), Excel.Application)
            crearExcel.Visible = True
            Dim ubExcel As Excel.Workbook = crearExcel.Workbooks.Open(Form1.TextBox1.Text)
            Dim hojaExcel As Excel.Worksheet = DirectCast(ubExcel.Worksheets("Hoja1"), Excel.Worksheet)

            hojaExcel.Range(ComboBox1.Text).Value = ""

            ubExcel.Close(SaveChanges:=True)
            crearExcel.Quit()

            'MsgBox("DATO ELIMINADO EXITOSAMENTE, ACTUALICE VENTANA PRINCIPAL", MsgBoxStyle.Information, "CONFIRMACIÓN")
            Form1.Show()

            ComboBox1.Text = ""
            ComboBox1.Focus()
        Else
            MsgBox("SELECCIONE CELDA", MsgBoxStyle.Critical, "ERROR")
        End If
    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For i As Integer = 2 To 100  '48576
            For j = 65 To 90
                'For x As Integer = 65 To 90
                Conversion = Chr(j) & i
                cadena = Conversion
                ComboBox1.Items.Add(cadena)
                'Next x
            Next j
            ComboBox1.Items.Add("----")
        Next i
    End Sub
End Class
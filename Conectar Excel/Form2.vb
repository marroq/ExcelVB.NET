Imports Microsoft.Office.Interop

Public Class Form2
    Dim j As Byte
    Dim cadena As String
    Dim conversion As String
    Dim dato As String
    Dim dato2 As String

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For i As Integer = 2 To 100  '48576
            For j = 65 To 90
                'For x As Integer = 65 To 90
                conversion = Chr(j) & i
                cadena = conversion
                ComboBox1.Items.Add(cadena)
                'Next x
            Next j
            ComboBox1.Items.Add("----")
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

            'MsgBox("DATO INGRESADO EXITOSAMENTE, ACTUALICE VENTANA PRINCIPAL", MsgBoxStyle.Information, "CONFIRMACIÓN")
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

'************OTROS************

'Dim ApplicationE As Excel.Application
'Dim LibroE As Excel._Workbook
'Dim LibrosE As Excel.Workbooks
'Dim HojasE As Excel.Sheets
'Dim HojaE As Excel._Worksheet
'Dim RangoE As Excel.Range

''Nueva instacia (Crear un libro nuevo)
'        ApplicationE = Form1.TextBox1.Text
'        LibrosE = ApplicationE.Workbooks
'        LibroE = LibrosE.Add
'        HojasE = LibroE.Worksheets
'        HojaE = HojasE(1)

''Obtener el rango en el que la celda de partida que tiene la direccion 'm_sStartingCell y sus dimensiones son m_iNumRows m_iNumCols x.
'        RangoE = HojaE.Range(ComboBox1.Text, Reflection.Missing.Value)
'        RangoE = RangoE.Resize(100, 10)

''Devolver el control de Excel al usuario
'        ApplicationE.Visible = True
'        ApplicationE.UserControl = True

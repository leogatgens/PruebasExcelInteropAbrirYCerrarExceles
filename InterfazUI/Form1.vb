Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim excelApp = New Microsoft.Office.Interop.Excel.Application

        Dim libro = excelApp.Workbooks.Add
        libro.Sheets(1).Cells(1, 1) = "Hola mundo"
        libro.SaveAs(Filename:="D:\ExcelTestAbrirYCerrrar\test2.xlsx")
        Label1.Text = "Correcto se guardo el excel"

        excelApp.Quit()

        libro = Nothing
        excelApp = Nothing

    End Sub
End Class

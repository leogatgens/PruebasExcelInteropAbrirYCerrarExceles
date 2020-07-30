Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub btCargueInsumos_Click(sender As Object, e As RibbonControlEventArgs) Handles btCargueInsumos.Click
        Dim excelApp = New Microsoft.Office.Interop.Excel.Application

        Dim libro = excelApp.Workbooks.Add
        libro.Sheets(1).Cells(1, 1) = "Hola mundo"
        libro.SaveAs(Filename:="D:\ExcelTestAbrirYCerrrar\test2.xlsx")


        excelApp.Quit()

        libro = Nothing
        excelApp = Nothing
    End Sub
End Class

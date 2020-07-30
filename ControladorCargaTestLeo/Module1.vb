Imports System.IO
Imports System.Threading
Imports Microsoft.Office.Interop

Module Module1
    Dim hilosProcesarPlantilla As New List(Of Thread)
    Sub Main()
        Dim losArchivos = GenerarListaArchivos()
        ' Process the list of .txt files found in the directory. '
        Dim fileName As String

        For Each fileName In losArchivos
            If (System.IO.File.Exists(fileName)) Then
                'Read File and Print Result if its true
                ProceseCargaPlantilla(fileName)
                Console.WriteLine(fileName)
            End If

        Next
        Console.ReadKey()

    End Sub

    Private Function GenerarListaArchivos() As List(Of String)
        Dim fileEntries As String() = Directory.GetFiles("D:\ExcelTestAbrirYCerrrar")

        Return fileEntries.ToList()

    End Function

    Private Sub ProceseCargaPlantilla(rutaArchivo As String)
        Dim hiloCargaPlantilla As Thread

        hiloCargaPlantilla = New Thread(Sub()

                                            Dim aplicacion As New Excel.Application With {
                                                                                              .Visible = True
                                                                                          }
                                            Dim libros As Excel.Workbooks = Nothing
                                            Dim libro As Excel.Workbook = Nothing
                                            Try
                                                Dim contraseña As String = String.Empty
                                                libros = aplicacion.Workbooks
                                                libro = libros.Open(rutaArchivo, UpdateLinks:=False)

                                                libro.Close(SaveChanges:=True)

                                                libros.Close()
                                                aplicacion.Quit()
                                                GC.Collect()
                                                GC.WaitForPendingFinalizers()
                                                Runtime.InteropServices.Marshal.ReleaseComObject(libro)
                                                Runtime.InteropServices.Marshal.ReleaseComObject(libros)
                                                Runtime.InteropServices.Marshal.ReleaseComObject(aplicacion)
                                                libro = Nothing
                                                libros = Nothing
                                                aplicacion = Nothing
                                            Catch ex As Runtime.InteropServices.COMException
                                                MsgBox("wtf Runtime.InteropServices.COMException" + ex.Message)

                                            Catch ex As Exception
                                                MsgBox("wtf Exception" + ex.Message)

                                            End Try

                                        End Sub)
        hiloCargaPlantilla.Start()
        hilosProcesarPlantilla.Add(hiloCargaPlantilla)
    End Sub

End Module

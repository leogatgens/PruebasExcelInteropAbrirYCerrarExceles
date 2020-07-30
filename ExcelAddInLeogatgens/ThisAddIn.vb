Imports System.Diagnostics
Imports System.Threading

Public Class ThisAddIn


    Private Sub ComplementoSEE_Startup() Handles Me.Startup



        AddHandler System.Windows.Forms.Application.ThreadException, AddressOf ApplicationThreadException



        AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf CurrentDomainUnhandledException
    End Sub

    Private Sub CurrentDomainUnhandledException(sender As Object, e As UnhandledExceptionEventArgs)

        MsgBox("CurrentDomainUnhandledException")
    End Sub

    Private Sub ApplicationThreadException(sender As Object, e As ThreadExceptionEventArgs)

        MsgBox("ApplicationThreadException")
    End Sub

    Private Sub WorkBook_Open(ByVal wb As Excel.Workbook) Handles Application.WorkbookOpen


        'Dim manejadorDeComponentes As ComponentesExcel.ManejadorDeInstancia
        'manejadorDeComponentes = ConstruyaManejadorDeComponentes()
        'Try
        '    VerifiqueSiEsArchivoDeReplicacion()
        '    If BC.Utilidades.Utilitarios.EsArchivoReplicacion Then
        '        CargueReplicacion()
        '    Else
        Dim excelApp As Excel.Application = Globals.ThisAddIn.Application
        Dim worksheet As Microsoft.Office.Interop.Excel.Workbook = excelApp.ActiveWorkbook
        Try


            Debug.WriteLine("OpenWork: inicio")
            Debug.WriteLine("wb.Name : " & wb.Name)

            'libro.Sheets(1).Cells(2, 1) = "Hola mundo"

            For i = 1 To 10000
                worksheet.Sheets(1).Cells(i, 1) = "Hola mundo"
            Next


            'worksheet.Close(SaveChanges:=True)
            If worksheet.Name.Contains("3") Then
                Throw New Exception("Se cayo esta vaa")
            End If


            ''    Private Microsoft.Office.Interop.Excel._Application excel;//excel application that Is used For creating excel workbooks(files)
            ''Private Microsoft.Office.Interop.Excel._Workbook workbook;//excel workbook(file)
            ''Private Microsoft.Office.Interop.Excel._Worksheet worksheet;

            ''Excel = New Microsoft.Office.Interop.Excel.Application();
            ''Workbook = Excel.Workbooks.Add();
            ''worksheet = Workbook.ActiveSheet;

            ''excelApp.Workbook.Close()
            'excelApp.Quit()

            'GC.Collect()
            'GC.WaitForPendingFinalizers()

            'System.Runtime.InteropServices.Marshal.FinalReleaseComObject(worksheet)
            ''System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp.Workbook)
            'System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp)

            'worksheet = Nothing
            ''Workbook = null;
            'excelApp = Nothing



        Catch ex As Exception
            MsgBox("Exception" & ex.Message)

            Throw


        End Try

        '        CarguePlantilla()
        '        BC.Utilidades.ManejadorDeMensajes.EscribaEvento("OpenWork termina")
        '    End If

        'Catch excepcion As System.ServiceModel.FaultException(Of Bccr.BaseFault)
        '    BC.Utilidades.ManejadorDeMensajes.EscribaEvento(excepcion.Message)
        '    Dim padre As NativeWindow = New NativeWindow()
        '    padre.AssignHandle(Process.GetCurrentProcess().MainWindowHandle)
        '    Utilidades.VisualizadorDeMensajes.MuestreMensajeDeErrorEnPantalla(excepcion.Message, padre)
        'Catch excepcion As System.ServiceModel.FaultException
        '    BC.Utilidades.ManejadorDeMensajes.EscribaEvento(excepcion.Message)
        '    Dim padre As NativeWindow = New NativeWindow()
        '    padre.AssignHandle(Process.GetCurrentProcess().MainWindowHandle)
        '    Utilidades.VisualizadorDeMensajes.MuestreMensajeDeErrorEnPantalla(excepcion.Message, padre)
        'Catch excepcion As BC.Utilidades.ClienteSEEException
        '    BC.Utilidades.ManejadorDeMensajes.EscribaEvento(excepcion.Message)
        '    Dim padre As NativeWindow = New NativeWindow()
        '    padre.AssignHandle(Process.GetCurrentProcess().MainWindowHandle)
        '    Utilidades.VisualizadorDeMensajes.MuestreMensajeDeErrorEnPantalla(excepcion.Message, padre)
        'Catch excepcion As System.Runtime.InteropServices.COMException
        '    BC.Utilidades.ManejadorDeMensajes.EscribaEvento("Se cae System.Runtime.InteropServices.COMException, no hacer nada ")

        'Catch excepcion As Exception
        '    BC.Utilidades.ManejadorDeMensajes.EscribaEvento(excepcion.Message)
        '    Dim padre As NativeWindow = New NativeWindow()
        '    padre.AssignHandle(Process.GetCurrentProcess().MainWindowHandle)
        '    Utilidades.VisualizadorDeMensajes.MuestreMensajeDeErrorEnPantalla(BC.My.Resources.Resources.CE0000, padre)
        'Finally
        '    manejadorDeComponentes.mensajes.EnciendaTodo()
        'End Try
    End Sub


    Private Sub ComplementoSEE_Shutdown() Handles Me.Shutdown
        'MsgBox("ComplementoSEE_Shutdown")
    End Sub

End Class

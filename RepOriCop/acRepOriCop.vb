Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel
#Region "Definiciones"
<Assembly: CLSCompliant(True)>
<Assembly: InternalsVisibleTo("OrionCopIU")>
<Assembly: InternalsVisibleTo("AdminOrionIU")>
<Assembly: InternalsVisibleTo("WinCom")>
<Assembly: InternalsVisibleTo("OrionCopL")>
<Assembly: InternalsVisibleTo("OriIntCon")>
#End Region
#Region "Enumeradores"
Friend Enum EnuReporteDef As Integer
    None = 0
    enuFactura
    enuFacturaEFac
    enuFacturaDscto
    enuFactImportada
    enuFactAutoMes
    enuExpFacsFechas
    enuRecCaja
    enuNotasAjuste
    enuNotasCon
    enuNotasDb
    enuNotaCr
    enuNotaDevAnt
    enuNotaReverCr
    enuAuxiliar
    enuAuxTer
    enuResumenMovCont
    enuCajaBancos
    enuRecCajaReversados
    enuInformeDiario
    enuCarteraPorCliente
    enuCarteraPorPredio
    enuCarteraPorServicio
    enuCarteraPorPredioAgr
    enuEdadCartera
    enuEdadCarteraDet
    enuEdadCarteraGrafico
    enuEstadoCuentas
    enuPrediosSector
    enuPrediosPropietario
    enuPropietariosXCP
    enuPropietariosXCP_Res
    enuCuotasAdminPropi
    enuItemsProgramaFact
    enuMovimCuenta
    enuMovimAntici
    enuDirTf
    enuCtaCobro
    enuCtaCobroDet
    enuEstadoCtaCli
    enuAnticiposPorAplicar
    enuAnticiposPorPredioAgru
    enuExpRecsCaja
    enuRCFechas
    enuRelDocs
    enuDirClientes
    enuDocsNoRegEFac
    enuCxCDetPorSer
    enuFacVivas
    enuValoresFacturados
    enuValoresFactTodos
End Enum

Friend Enum EnuTipoRepEdadCartera As Integer
    None = 0
    enuDetallado
    enuResumido
    enuGrafico
End Enum
#End Region
Friend Class ClsExportToExcell
    Private ReadOnly MappExcel As Microsoft.Office.Interop.Excel.Application
    Private MwbLibro As Workbook = Nothing
    Private MwsHoja As Worksheet = Nothing
    Private MdgOrigenDatos As DataGrid = Nothing
    Private MstrNombreArchivo As String = String.Empty
    Friend Sub New()
        MappExcel = New Microsoft.Office.Interop.Excel.Application()
    End Sub
    ''' <summary>
    ''' Exporta el contenido de un datagrid cuyo datacontext es un datatable a un archivo de Excel
    ''' </summary>
    ''' <param name="astrTitulo">Es el titulo que se pondra ebcima de todas las columnas de la hoja de excel</param>
    ''' <param name="astrArchivoExcel">Nombre del archivo de excel incluida su ruta.</param>
    ''' <param name="adgOrigenDatos">El datagrid del cual se exportaran los datos.</param>
    ''' <param name="adtbOrigenDatos">El datatable que es el datacontext del datagrid.</param>
    ''' <param name="astrMapeoColumnas">Es un array que contiene el mapeo del nombre de las columnas
    ''' en el datagrid con el nombre de las columnas en el datatable.</param>
    ''' <remarks></remarks>
    Friend Function FblnExportToExcel(astrTitulo As String, astrArchivoExcel As String,
            adgOrigenDatos As DataGrid, adtbOrigenDatos As System.Data.DataTable,
            astrMapeoColumnas() As String) As Boolean
        Dim lblnNoHayError As Boolean = False
        Dim lstrMens As String = String.Empty
        Try
            MstrNombreArchivo = astrArchivoExcel
            Dim lstrNombreHoja = ClsPanoramaDat.FstrObtengaNombre(astrArchivoExcel)
            MdgOrigenDatos = adgOrigenDatos
            SCreeLibro(lstrNombreHoja)
            SInserteEncabezado(astrTitulo)
            SInserteDatos(adtbOrigenDatos, astrMapeoColumnas)
            STermine()
            lblnNoHayError = True
        Catch ex As Exception
            lstrMens = "No fue posible crear el archivo de Excel debido posiblemente a problemas " &
                    "con la licencia de Office. Por favor, verifique que su licencia esté activa " &
                    "y sea válida."
        Finally
            If Not lblnNoHayError Then
                MsgBox(lstrMens, vbOKOnly, "Información")
            End If
        End Try
        Return lblnNoHayError
    End Function
    Friend Function FblnExportoToExcel(astrArchivoExcel As String,
                adtbOrigenDatos As System.Data.DataTable) As Boolean
        Dim lblnNoHayError = False, lstrMens = String.Empty
        Try
            MstrNombreArchivo = astrArchivoExcel
            If My.Computer.FileSystem.FileExists(MstrNombreArchivo & ".xlsx") Then
                My.Computer.FileSystem.DeleteFile(MstrNombreArchivo & ".xlsx")
            End If
            Dim lstrNombreHoja = ClsPanoramaDat.FstrObtengaNombre(astrArchivoExcel)
            SCreeLibro(lstrNombreHoja)
            SInserteEncabezado(adtbOrigenDatos, lstrNombreHoja)
            SInserteDatos(adtbOrigenDatos, False, False)
            STermine()
            lblnNoHayError = True
        Catch ex As System.IO.IOException
            lstrMens = "No fue posible exportar a Excel porque el archivo está en uso!"
        Catch ex As Exception
            lstrMens = "No fue posible exportar a Excel debido posiblemente a problemas " &
                    "con la licencia de Office. Por favor, verifique que su licencia esté activa " &
                    "y sea válida."
        Finally
            If Not lblnNoHayError Then
                MsgBox(lstrMens, vbOKOnly, "Información")
            End If
        End Try
        Return lblnNoHayError
    End Function
    Friend Sub SExporteToExcel(astrArchivoExcel As String, adtbOrigenDatos As System.Data.DataTable,
                astrNombresColumnas As String())
        Dim lblnNoHayError = False
        Dim lstrMens As String = String.Empty
        Try
            MstrNombreArchivo = astrArchivoExcel
            Dim lstrNombreHoja = ClsPanoramaDat.FstrObtengaNombre(astrArchivoExcel)
            SCreeLibro(lstrNombreHoja)
            SInserteEncabezado(astrNombresColumnas, lstrNombreHoja)
            SInserteDatos(adtbOrigenDatos, False, False)
            STermine()
            lblnNoHayError = True
        Catch ex As Exception
            lstrMens = "No fue posible crear el archivo de Excel debido posiblemente a problemas " &
                    "con la licencia de Office. Por favor, verifique que su licencia esté activa " &
                    "y sea válida."
        Finally
            If Not lblnNoHayError Then
                MsgBox(lstrMens, vbOKOnly, "Información")
            End If
        End Try
    End Sub
    Friend Sub SExporteToExcel(astrArchivoExcel As String, adtbOrigenDatos As System.Data.DataTable,
            astrNombresColumnas As String(), ablnEdadCartera As Boolean, ablnDetallado As Boolean)
        Dim lblnNoHayError = False
        Dim lstrMens As String = String.Empty
        Try
            MstrNombreArchivo = astrArchivoExcel
            Dim lstrNombreHoja = ClsPanoramaDat.FstrObtengaNombre(astrArchivoExcel)
            SCreeLibro(lstrNombreHoja)
            SInserteEncabezado(astrNombresColumnas, lstrNombreHoja)
            SInserteDatos(adtbOrigenDatos, ablnEdadCartera, ablnDetallado)
            STermine()
            lblnNoHayError = True
        Catch ex As Exception
            lstrMens = "No fue posible crear el archivo de Excel debido posiblemente a problemas " &
                    "con la licencia de Office. Por favor, verifique que su licencia esté activa " &
                    "y sea válida."
        Finally
            If Not lblnNoHayError Then
                MsgBox(lstrMens, vbOKOnly, "Información") 'JSV Muestra el mensaje en pantalla
            End If
        End Try
    End Sub
    Friend Sub SCreeLibro(astrNombreHoja As String)
        MwbLibro = MappExcel.Workbooks.Add()
        MwsHoja = MwbLibro.Worksheets.Add()
        MwsHoja.Name = astrNombreHoja
    End Sub
    Private Sub SInserteEncabezado(astrTitulo As String)
        Dim lentIndice = MdgOrigenDatos.Columns.Count
        If lentIndice > 10 Then lentIndice = 10
        Dim lstrColumnaFinal = FstrLetraDeIndice(lentIndice)
        Dim lstrRango1 = "A1:" & lstrColumnaFinal & "1"
        Dim lstrRango2 = "A2:" & lstrColumnaFinal & "2"
        MwsHoja.Range(lstrRango1).Merge()
        MwsHoja.Range(lstrRango1).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
        MwsHoja.Range(lstrRango1).Value = GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual.ObjNombreCentroUtilStr.ObjValorPro
        MwsHoja.Range(lstrRango2).Merge()
        MwsHoja.Range(lstrRango2).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
        MwsHoja.Range(lstrRango2).Value = astrTitulo
        For i = 0 To MdgOrigenDatos.Columns.Count - 1
            Dim lstrEncabezado = MdgOrigenDatos.Columns(i).Header
            MwsHoja.Cells(4, i + 1) = lstrEncabezado
        Next
    End Sub
    Private Sub SInserteEncabezado(adtbOrigenDatos As System.Data.DataTable,
            astrTitulo As String)
        MwsHoja.Cells(1, 1) = GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual.ObjNombreCentroUtilStr.ObjValorPro
        MwsHoja.Cells(2, 1) = astrTitulo
        For i = 0 To adtbOrigenDatos.Columns.Count - 1
            Dim lstrEncabezado = adtbOrigenDatos.Columns(i).ColumnName
            MwsHoja.Cells(4, i + 1) = lstrEncabezado
        Next
    End Sub
    Private Sub SInserteEncabezado(astrNombresColumnas() As String, astrTitulo As String)
        MwsHoja.Cells(1, 1) = GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual.ObjNombreCentroUtilStr.ObjValorPro
        MwsHoja.Cells(2, 1) = astrTitulo
        For i = 0 To astrNombresColumnas.Length - 1
            Dim lstrEncabezado = astrNombresColumnas(i)
            MwsHoja.Cells(4, i + 1) = lstrEncabezado
        Next
    End Sub
    Friend Sub SInserteEncabezadoHoja(astrNombresColumnas As String(),
                aentFilaEncabezado As Integer)
        Dim lentIndice = astrNombresColumnas.Length
        If aentFilaEncabezado > 1 Then
            If lentIndice > 10 Then lentIndice = 10
            Dim lstrColumnaFinal = FstrLetraDeIndice(lentIndice)
            Dim lstrRango1 = "A1:" & lstrColumnaFinal & "1"
            Dim lstrRango2 = "A2:" & lstrColumnaFinal & "2"
            Dim lstrRango3 = "A3:" & lstrColumnaFinal & "3"
            Dim lstrRango4 = "A4:" & lstrColumnaFinal & "4"
            MwsHoja.Range(lstrRango1).Merge()
            MwsHoja.Range(lstrRango1).HorizontalAlignment =
                Microsoft.Office.Interop.Excel.Constants.xlLeft
            MwsHoja.Range(lstrRango2).Merge()
            MwsHoja.Range(lstrRango2).HorizontalAlignment =
                    Microsoft.Office.Interop.Excel.Constants.xlLeft
            MwsHoja.Range(lstrRango3).Merge()
            MwsHoja.Range(lstrRango3).HorizontalAlignment =
                    Microsoft.Office.Interop.Excel.Constants.xlLeft
            MwsHoja.Range(lstrRango4).Merge()
            MwsHoja.Range(lstrRango4).HorizontalAlignment =
                    Microsoft.Office.Interop.Excel.Constants.xlLeft
            MwsHoja.Range(lstrRango1).Value = GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual.ObjNombreCentroUtilStr.ObjValorPro
        End If
        For i = 0 To astrNombresColumnas.Length - 1
            MwsHoja.Cells(aentFilaEncabezado, i + 1) = astrNombresColumnas(i)
        Next
    End Sub
    Private Sub SInserteDatos(adtbOrigenDatos As System.Data.DataTable,
            astrMapeoCols() As String)
        Dim lstrColumnaFinal = FstrLetraDeIndice(MdgOrigenDatos.Columns.Count)
        Dim lobjDato As Object
        Dim j = 5
        For Each ldrwDato As DataRow In adtbOrigenDatos.Rows
            For i = 0 To MdgOrigenDatos.Columns.Count - 1
                Dim lstrNomColDtb = FstrNombreColumnaDtb(astrMapeoCols, MdgOrigenDatos.Columns(i).Header)
                lobjDato = ldrwDato(lstrNomColDtb)
                MwsHoja.Cells(j, i + 1) = lobjDato
            Next
            j += 1
        Next
        Dim lstrRango = "A4:" & lstrColumnaFinal & j.ToString
        Dim lobjRango As Microsoft.Office.Interop.Excel.Range = MwsHoja.Range(lstrRango)
        lobjRango.Select()
        lobjRango.Columns.AutoFit()
        lobjRango = MwsHoja.Range("A5:A5")
        lobjRango.Select()
    End Sub
    Private Sub SInserteDatos(adtbOrigenDatos As System.Data.DataTable,
            ablnEdadCartera As Boolean, ablnDetallado As Boolean)
        Dim lstrColumnaFinal = FstrLetraDeIndice(adtbOrigenDatos.Columns.Count)
        Dim lobjDato As Object
        Dim ldecValor As Decimal, ldecTotal As Decimal, lentColIni = 0, lentColFin = 0
        Dim ldecGranTotal(7) As Decimal
        Dim j = 5
        If ablnEdadCartera Then
            If ablnDetallado Then
                lentColIni = 6
                lentColFin = 11
            Else
                lentColIni = 3
                lentColFin = 8
            End If
        End If
        For Each ldrwDato As DataRow In adtbOrigenDatos.Rows
            ldecTotal = 0
            For i = 0 To adtbOrigenDatos.Columns.Count - 1
                lobjDato = ldrwDato(i)
                If ablnEdadCartera Then
                    If i >= lentColIni And i <= lentColFin Then
                        ldecValor = ClsPanorama.FobjValorCampo(lobjDato, EnuTipoValor.enuDecimal)
                        ldecGranTotal(i - lentColIni) += ldecValor
                        If Not ablnDetallado Then
                            MwsHoja.Cells(j, i + 1) = ldecValor
                            ldecTotal += ldecValor
                            MwsHoja.Cells(j, i + 2) = ldecTotal
                        Else
                            MwsHoja.Cells(j, i + 1) = ldecValor
                            ldecTotal += ldecValor
                        End If
                    ElseIf i = 12 Then
                        MwsHoja.Cells(j, i + 1) = ldecTotal
                        MwsHoja.Cells(j, i + 2) = lobjDato
                    Else
                        MwsHoja.Cells(j, i + 1) = lobjDato
                    End If
                Else
                    MwsHoja.Cells(j, i + 1) = lobjDato
                End If
            Next
            j += 1
        Next
        For i = 1 To 6
            ldecGranTotal(6) += ldecGranTotal(i - 1)
            MwsHoja.Cells(j, lentColIni + i) = ldecGranTotal(i - 1)
        Next
        MwsHoja.Cells(j, lentColIni + 7) = ldecGranTotal(6)
        Dim lstrRango = "A4:" & lstrColumnaFinal & j.ToString
        Dim lobjRango As Microsoft.Office.Interop.Excel.Range = MwsHoja.Range(lstrRango)
        lobjRango.Select()
        lobjRango.Columns.AutoFit()
        lobjRango = MwsHoja.Range("A5:A5")
        lobjRango.Select()
    End Sub
    Friend Sub SInserteDatos(aentOrdinal As Integer, aobjDatos As Object(),
            ablnConEncabezado As Boolean)
        Dim lstrColumnaFinal = FstrLetraDeIndice(aobjDatos.Length)
        Dim lentIdFila = aentOrdinal, lentIdCol = 0
        If ablnConEncabezado Then
            lentIdFila += 5
        End If
        For Each lobjDato As Object In aobjDatos
            lentIdCol += 1
            MwsHoja.Cells(lentIdFila, lentIdCol) = lobjDato
        Next
        Dim lstrRango = "A5:" & lstrColumnaFinal & lentIdFila.ToString
        Dim lobjRango As Microsoft.Office.Interop.Excel.Range = MwsHoja.Range(lstrRango)
        lobjRango.Select()
        lobjRango.Columns.AutoFit()
        lobjRango = MwsHoja.Range("A5:A5")
        lobjRango.Select()
    End Sub
    Private Shared Function FstrNombreColumnaDtb(astrMapeoCols() As String,
            astrNombreColDataGrid As String) As String
        Dim lstrNomColDtb = String.Empty
        For Each lstrMapa As String In astrMapeoCols
            Dim lstrColumnas() = lstrMapa.Split("=")
            If lstrColumnas(0).ToUpper.Trim = astrNombreColDataGrid.ToUpper.Trim Then
                lstrNomColDtb = lstrColumnas(1).Trim
                Exit For
            End If
        Next
        Return lstrNomColDtb
    End Function
    Private Sub STermine()
        If Not IsNothing(MappExcel) Then
            If Not IsNothing(MappExcel.ActiveWorkbook) Then
                Dim ldblResul As Double = 0, lblnNoHayError = False, lstrMens = String.Empty
                Try
                    MappExcel.ActiveWorkbook.SaveAs(MstrNombreArchivo)
                    lblnNoHayError = True
                Catch ex As Exception
                    ldblResul = ex.HResult
                    lstrMens = ex.Message
                Finally
                    If Not IsNothing(MwbLibro) Then
                        If Not IsNothing(MwsHoja) Then
                            Marshal.ReleaseComObject(MwsHoja)
                            MwbLibro.Close(False, Type.Missing, Type.Missing)
                            Marshal.ReleaseComObject(MwbLibro)
                        End If
                    End If
                    MappExcel.Quit()
                    Marshal.ReleaseComObject(MappExcel)
                    If Not lblnNoHayError Then
                        If Not ldblResul = -2146827284 Then
                            Throw New Exception(lstrMens)
                        End If
                    End If
                End Try
            End If
        End If
    End Sub
    Friend Sub STermine(astrNombreArchivo As String)
        Dim lblnNoHayError = False, ldblResul As Double = 0
        MstrNombreArchivo = astrNombreArchivo
        If Not IsNothing(MappExcel) Then
            If Not IsNothing(MappExcel.ActiveWorkbook) Then
                Try
                    MappExcel.ActiveWorkbook.SaveAs(MstrNombreArchivo)
                    lblnNoHayError = True
                Catch ex As Exception
                    ldblResul = ex.HResult
                Finally
                    If Not IsNothing(MwbLibro) Then
                        If Not IsNothing(MwsHoja) Then
                            Marshal.ReleaseComObject(MwsHoja)
                            MwbLibro.Close(False, Type.Missing, Type.Missing)
                            Marshal.ReleaseComObject(MwbLibro)
                        End If
                    End If
                    MappExcel.Quit()
                    Marshal.ReleaseComObject(MappExcel)
                    If Not lblnNoHayError Then
                        If Not ldblResul = -2146827284 Then
                            Throw New Exception
                        End If
                    End If
                End Try
            End If
        End If
    End Sub
    Public Sub SAbraExcel()
        Dim lstrNomArch = MstrNombreArchivo & ".xlsx"
        Dim lblnExiste = My.Computer.FileSystem.FileExists(lstrNomArch)
        If Not lblnExiste Then
            lstrNomArch = MstrNombreArchivo & ".xls"
            lblnExiste = My.Computer.FileSystem.FileExists(lstrNomArch)
        End If
        If lblnExiste Then
            Dim lappExcel = New Microsoft.Office.Interop.Excel.Application
            lappExcel.Workbooks.Open(MstrNombreArchivo)
            lappExcel.Visible = True
        Else
            Throw New ErrorInesperadoPanLException("El Archivo de Excel No Existe")
        End If
    End Sub
End Class
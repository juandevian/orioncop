Imports System.IO
Imports System.Text
Friend Class ClsInterfazContaColon
    Inherits ClsCBInterfazContableOri
#Region "Definiciones"
    'Constantes
    'Variables
    Private MstrNombreArchivo As String = String.Empty
    Private MblnLeidoUltimoRegistro As Boolean = False
    Private MdtmFechaFinInterfazAnterior As Date = GCDTMFECHANULA
    Private ReadOnly MstbInterfaz As New StringBuilder
#End Region
#Region "Constructores"
    Friend Sub New(aobjRegistro As Object)
        MyBase.New(aobjRegistro)
    End Sub
#End Region
#Region "Propiedades"
    '
#End Region
#Region "Interfaz"
    Friend Overrides Sub SGenerereInterfazContable(ablnFinMes As Boolean, ByRef astrMens As String)
        Dim lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            If FblnExisteExcel() Then
                If Not FblnEstanTodosOkInterfaz() Then
                    Throw New ValorArgumentoInvalidoException("Hay argumentos no validos")
                End If
                Dim lstrFechaDesde = ClsPanorama.FstrFechayyyymmdd(DtmFechaDesde)
                Dim lstrFechaHasta = ClsPanorama.FstrFechayyyymmdd(DtmFechaHasta)
                If ablnFinMes Then
                    MstrNombreArchivo = "InterfazColon_FM_" & lstrFechaDesde & "_" & lstrFechaHasta
                Else
                    MstrNombreArchivo = "InterfazColon_" & lstrFechaDesde & "_" & lstrFechaHasta
                End If
                HstrArchivoSalidaInterfaz = GstrTrayInterfContable & "\" & MstrNombreArchivo
                If My.Computer.FileSystem.FileExists(HstrArchivoSalidaInterfaz & ".xlsx") Then
                    My.Computer.FileSystem.DeleteFile(HstrArchivoSalidaInterfaz & ".xlsx")
                End If
                SGenereIntColon(astrMens)
                SExporteAExcel(HstrArchivoSalidaInterfaz)
                MdtmFechaFinInterfazAnterior = GCDTMFECHANULA
            Else
                astrMens = "No se puede generar debido a que no está instalado Excel"
            End If
            lblnNoHayError = True
        Catch ex As ErrorInesperadoPanLException
            Throw
        Catch ex As ArgumentException
            Throw
        Catch ex As PathTooLongException
            Throw
        Catch ex As DirectoryNotFoundException
            Throw
        Catch ex As System.IO.IOException
            astrMens = "No se generó la Interfaz debido a que el archivo de Excel está abierto!"
        Catch ex As NotSupportedException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                GobjPanDat.SControleProcesoObj(False)
                If String.IsNullOrEmpty(astrMens) Then
                    astrMens = "La Interfaz Contable se generó exitosamente!"
                End If
            Else
                GobjPanDat.SControleProcesoObj(False, True)
                MsgBox(astrMens, vbOKOnly, "Información")
            End If
        End Try
    End Sub
    Private Sub SGenereIntColon(ByRef astrMens As String)
        Dim ldtbMovi = FdtbMovimiento()
        Try
            Using lswInterfaz = ClsPanorama.FswStreamWriter(HstrArchivoSalidaInterfaz & ".csv")
                SEscribaEncabezado(lswInterfaz)
                SProceseDoc(ldtbMovi, EnuIdDocumentoDef.EnuFacturaVenta, lswInterfaz)
                SProceseDoc(ldtbMovi, EnuIdDocumentoDef.EnuNotaIntMora, lswInterfaz)
                SProceseRecibosCaja(lswInterfaz)
                SProceseDoc(ldtbMovi, EnuIdDocumentoDef.EnuNotaAplicacionAnt, lswInterfaz)
                SProceseDoc(ldtbMovi, EnuIdDocumentoDef.EnuNotaCr, lswInterfaz)
                SProceseNotaR(EnuIdDocumentoDef.EnuNotaReintegroAnt, lswInterfaz)
                SProceseDoc(ldtbMovi, EnuIdDocumentoDef.EnuNotaReversaCr, lswInterfaz)
                SProceseNotaR(EnuIdDocumentoDef.EnuNotaReversaCr, lswInterfaz)
                SProceseNotasAjuste(lswInterfaz)
            End Using
        Catch ex As Exception
            astrMens = ex.Message
        Finally
            If Not String.IsNullOrEmpty(astrMens) Then
                MsgBox(astrMens, vbOKOnly, "Información")
            End If
        End Try
    End Sub
    Private Sub SProceseDoc(adtbMovi As DataTable, aenuIdDocumento As EnuIdDocumentoDef,
            aswInterfaz As StreamWriter)
        Dim ldecValor As Decimal
        Dim lstrPreFac As String, lentIdFac As Integer
        Dim lstrNroFac As String
        Dim lstrNroFacAnt = String.Empty
        Dim lstrTipoDoc = GobjParametros.FstrTipoDoc(aenuIdDocumento)
        Dim lstrFiltro = ClsIdTipoDocOrigenByt.SstrNombreCampoBd & " = " &
                aenuIdDocumento
        Dim ldrwMvimDoc As DataRow() = adtbMovi.Select(lstrFiltro)
        For Each ldrwNov As DataRow In ldrwMvimDoc
            lstrPreFac = CType(ClsPanorama.FobjValorCampo(ldrwNov(
                        ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString), String)
            If IsNothing(lstrPreFac) Then lstrPreFac = String.Empty
            lentIdFac = ClsPanorama.FobjValorCampo(ldrwNov(
                        ClsIdDocOrigenEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
            If aenuIdDocumento = EnuIdDocumentoDef.EnuFacturaVenta Then
                lstrNroFac = ClsPanorama.FstrNumeroDcto(lstrPreFac, lentIdFac)
                If lstrNroFacAnt <> lstrNroFac Then
                    SProceseFactura(adtbMovi, lstrPreFac, lentIdFac, aswInterfaz)
                    lstrNroFacAnt = lstrNroFac
                End If
            Else
                ldecValor = ClsPanorama.FobjValorCampo(ldrwNov(
                        ClsValor_NovDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
                SEscribaMovDoc(ldrwNov, lstrTipoDoc, lstrPreFac, lentIdFac, ldecValor,
                        True, aswInterfaz)
                SEscribaMovDoc(ldrwNov, lstrTipoDoc, lstrPreFac, lentIdFac, ldecValor,
                        False, aswInterfaz)
            End If
        Next
    End Sub
    Private Sub SProceseFactura(adtbMovi As DataTable, astrPrefFac As String,
            aentIdFac As Integer, aswInterfaz As StreamWriter)
        Dim lstrFiltro = ClsIdTipoDocOrigenByt.SstrNombreCampoBd & " = " &
                EnuIdDocumentoDef.EnuFacturaVenta & " AND " &
                ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd & " = '" & astrPrefFac &
                "' AND " & ClsIdDocOrigenEnt.SstrNombreCampoBd & " = " & aentIdFac
        Dim ldrwNovsFra As DataRow() = adtbMovi.Select(lstrFiltro), ldrwNov As DataRow
        Dim ldecValorFra = 0D, ldecValor As Decimal, lstrIdCtaDb = String.Empty, i = 0
        Dim lshrIdItemFac As Short
        Dim lstrIdTerceroCtaCr As String
        Dim lstrTipoDoc = GobjParametros.FstrTipoDoc(EnuIdDocumentoDef.EnuFacturaVenta)
        Dim ldrwNovDbFra As DataRow = Nothing
        Dim lenuTipoNovDbFra As EnuTipoNov
        Dim lobjFactura As New ClsFactura()
        Dim lobjValorLlave As Object()
        Do
            ldrwNov = ldrwNovsFra(i)
            lenuTipoNovDbFra = ClsPanorama.FobjValorCampo(ldrwNov(
                    ClsIdTipoNovedadByt.SstrNombreCampoBd), EnuTipoValor.enuByte)
            If lstrIdCtaDb <> ClsPanorama.FobjValorCampo(ldrwNov(
                    ClsIdCuentaDb_NovStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString) Then
                If Not String.IsNullOrEmpty(lstrIdCtaDb) Then
                    SEscribaMovDoc(ldrwNovsFra(i - 1), lstrTipoDoc, astrPrefFac,
                            aentIdFac, ldecValorFra, True, aswInterfaz)
                End If
                lstrIdCtaDb = ClsPanorama.FobjValorCampo(ldrwNov(
                        ClsIdCuentaDb_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
                ldecValorFra = ClsPanorama.FobjValorCampo(ldrwNov(
                         ClsValor_NovDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
            Else
                ldecValorFra += ClsPanorama.FobjValorCampo(ldrwNov(
                        ClsValor_NovDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
            End If
            If lenuTipoNovDbFra = EnuTipoNov.EnuDbCap Then
                ldrwNovDbFra = ldrwNov
            End If
            If i = ldrwNovsFra.Length - 1 Then
                SEscribaMovDoc(ldrwNovDbFra, lstrTipoDoc, astrPrefFac, aentIdFac,
                        ldecValorFra, True, aswInterfaz)
            End If
            i += 1
        Loop While i < ldrwNovsFra.Length
        For Each ldrwNovFac As DataRow In ldrwNovsFra
            ldecValor = ClsPanorama.FobjValorCampo(ldrwNovFac(
                    ClsValor_NovDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
            lshrIdItemFac = ClsPanorama.FobjValorCampo(ldrwNovFac(
                    ClsIdItemFacturaShr.SstrNombreCampoBd), EnuTipoValor.enuShort)
            lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, astrPrefFac, aentIdFac}
            lobjFactura.SAbra(lobjValorLlave)
            lstrIdTerceroCtaCr = lobjFactura.FstrIdTerceroCtaCr(lshrIdItemFac)
            If String.IsNullOrEmpty(lstrIdTerceroCtaCr) Then
                SEscribaMovDoc(ldrwNovFac, lstrTipoDoc, astrPrefFac, aentIdFac,
                        ldecValor, False, aswInterfaz)
            Else
                SEscribaMovCrFac(ldrwNovFac, lstrTipoDoc, astrPrefFac, aentIdFac,
                        ldecValor, lstrIdTerceroCtaCr, aswInterfaz)
            End If
        Next
    End Sub
    Private Sub SProceseRecibosCaja(aswInterfaz As StreamWriter)
        Dim lstrTipoDoc = GobjParametros.FstrTipoDoc(EnuIdDocumentoDef.EnuReciboCaja)
        Dim lstrPrefRC = String.Empty, lentIdRC = 0, lstrPrefRCNov As String
        Dim lentIdRCNov As Integer, lstrFiltro As String, lstrIdCuentaDb = String.Empty
        Dim lstrIdCuentaDbNov As String, lstrIdRC As String
        Dim ldrwMovRC As DataRow()
        Dim ldtbMovRC = FdtbMoviRC()
        For Each ldrwNov As DataRow In ldtbMovRC.Rows
            lstrPrefRCNov = ClsPanorama.FobjValorCampo(ldrwNov(
                    ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            lentIdRCNov = ClsPanorama.FobjValorCampo(ldrwNov(
                    ClsIdDocOrigenEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
            lstrIdRC = lentIdRCNov.ToString().PadLeft(7, "0")
            lstrIdCuentaDbNov = ClsPanorama.FobjValorCampo(ldrwNov(
                ClsIdCuentaDb_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString)
            If lstrPrefRC <> lstrPrefRCNov OrElse lentIdRC <> lentIdRCNov OrElse
                    lstrIdCuentaDb <> lstrIdCuentaDbNov Then
                lstrFiltro = ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd & " = '" &
                        lstrPrefRCNov & "' AND " & ClsIdDocOrigenEnt.SstrNombreCampoBd &
                        " = " & lentIdRCNov & " AND " &
                        ClsIdCuentaDb_NovStr.SstrNombreCampoBd & " = '" &
                        lstrIdCuentaDbNov & "'"
                ldrwMovRC = ldtbMovRC.Select(lstrFiltro)
                SProceseRecCaja(ldrwMovRC, lstrTipoDoc, aswInterfaz)
                lstrPrefRC = lstrPrefRCNov
                lentIdRC = lentIdRCNov
                lstrIdCuentaDb = lstrIdCuentaDbNov
            End If
        Next
    End Sub
    Private Sub SProceseRecCaja(adrwMoviRC As DataRow(), astrTipoDoc As String,
            aswInterfaz As StreamWriter)
        Dim ldecValor = 0D, ldecValorNov As Decimal, lstrPreFac As String
        Dim lentIdFac As Integer
        ' Calcular valor debito 
        For Each ldrwMovNov As DataRow In adrwMoviRC
            ldecValorNov = ClsPanorama.FobjValorCampo(ldrwMovNov(
                    ClsValor_NovDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
            ldecValor += ldecValorNov
        Next
        SEscribaMovDoc(adrwMoviRC(0), astrTipoDoc, "", 0, ldecValor, True, aswInterfaz)
        For Each ldrwMovNov As DataRow In adrwMoviRC
            lstrPreFac = CType(ClsPanorama.FobjValorCampo(ldrwMovNov(
                        ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString), String)
            If IsNothing(lstrPreFac) Then lstrPreFac = String.Empty
            lentIdFac = ClsPanorama.FobjValorCampo(ldrwMovNov(
                        ClsIdDocOrigenEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
            ldecValorNov = ClsPanorama.FobjValorCampo(ldrwMovNov(
                    ClsValor_NovDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
            SEscribaMovDoc(ldrwMovNov, astrTipoDoc, lstrPreFac, lentIdFac, ldecValorNov,
                    False, aswInterfaz)
        Next
    End Sub
    Private Sub SProceseNotaR(aenuIdDocumento As EnuIdDocumentoDef, aswInterfaz As StreamWriter)
        Dim ldecValor As Decimal
        Dim lstrTipoDoc = GobjParametros.FstrTipoDoc(aenuIdDocumento)
        Dim ldtbMovAntR = FdtbMoviAnt(aenuIdDocumento)
        For Each ldrwNov As DataRow In ldtbMovAntR.Rows
            ldecValor = ClsPanorama.FobjValorCampo(ldrwNov(
                    ClsValor_NovDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
            SEscribaMovDoc(ldrwNov, lstrTipoDoc, "", 0, ldecValor, True, aswInterfaz)
            SEscribaMovDoc(ldrwNov, lstrTipoDoc, "", 0, ldecValor, False, aswInterfaz)
        Next
    End Sub
    Private Sub SProceseNotasAjuste(aswInterfaz As StreamWriter)
        Dim lstrTipoDoc = GobjParametros.FstrTipoDoc(EnuIdDocumentoDef.EnuNotaAjuste)
        Dim ldecValor As Decimal
        Dim ldtbMovNA = FdtbMoviNotaAjuste()
        For Each ldrwNov As DataRow In ldtbMovNA.Rows
            ldecValor = ClsPanorama.FobjValorCampo(ldrwNov(
                    ClsValor_NovDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
            SEscribaMovDoc(ldrwNov, lstrTipoDoc, "", 0, ldecValor, True, aswInterfaz)
            SEscribaMovDoc(ldrwNov, lstrTipoDoc, "", 0, ldecValor, False, aswInterfaz)
        Next
    End Sub
#Region "Archivo Interfaz"
    Private Sub SEscribaEncabezado(aswInterfaz As StreamWriter)
        With MstbInterfaz
            .Clear().Append("cuenta").Append(CHSTRCOMA).Append("tercero").Append(CHSTRCOMA)
            .Append("dctos").Append(CHSTRCOMA).Append("prefijo").Append(CHSTRCOMA)
            .Append("documento").Append(CHSTRCOMA).Append("vendedor").Append(CHSTRCOMA)
            .Append("concepto").Append(CHSTRCOMA).Append("fecha").Append(CHSTRCOMA)
            .Append("cheque").Append(CHSTRCOMA).Append("bancoche").Append(CHSTRCOMA)
            .Append("fechache").Append(CHSTRCOMA).Append("valor").Append(CHSTRCOMA)
            .Append("mvto").Append(CHSTRCOMA).Append("cencos").Append(CHSTRCOMA)
            .Append("prefaf").Append(CHSTRCOMA).Append("documento").Append(CHSTRCOMA)
            .Append("base").Append(CHSTRCOMA).Append("fechatran").Append(CHSTRCOMA)
            .Append("horatran").Append(CHSTRCOMA).Append("usuario").Append(CHSTRCOMA)
            .Append("fechamod").Append(CHSTRCOMA).Append("horamod").Append(CHSTRCOMA)
            .Append("usuariomod").Append(CHSTRCOMA).Append("equipo").Append(CHSTRCOMA)
            .Append("vencimient").Append(CHSTRCOMA).Append("porreten").Append(CHSTRCOMA)
            .Append("inmueble").Append(CHSTRCOMA).Append("consoli").Append(CHSTRCOMA)
            .Append("zona").Append(CHSTRCOMA).Append("ruta").Append(CHSTRCOMA)
            .Append("cuota").Append(CHSTRCOMA).Append("niif").Append(CHSTRCOMA)
            .Append("cobrador")
        End With
        aswInterfaz.WriteLine(MstbInterfaz.ToString())
    End Sub
    Private Sub SEscribaMovDoc(adrwMov As DataRow, astrTipoDoc As String,
            astrPrefFac As String, aentIdFact As Integer, adecValor As Decimal,
            ablnDb As Boolean, aswInterfaz As StreamWriter)
        Dim lstrIdFac As String
        Dim lstrPrefDoc = ClsPanorama.FobjValorCampo(adrwMov(
                ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString)
        Dim lentIdDoc = ClsPanorama.FobjValorCampo(adrwMov(
                ClsIdDocOrigenEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
        Dim lstrIdDoc = "'" & lentIdDoc.ToString.PadLeft(7, "0")
        If aentIdFact = 0 Then
            lstrIdFac = ""
        Else
            lstrIdFac = "'" & aentIdFact.ToString().PadLeft(7, "0")
        End If
        Dim ldecBase = ClsPanorama.FobjValorCampo(adrwMov(ClsBaseDec.SstrNombreCampoBd),
                EnuTipoValor.enuDecimal)
        Dim lstrIdPredAgr As String = ClsPanorama.FobjValorCampo(adrwMov(
                ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString)
        Dim lstrCuentaCont As String
        Dim lstrTipoCuenta As String
        If ablnDb Then
            lstrTipoCuenta = "D"
            lstrCuentaCont = ClsPanorama.FobjValorCampo(adrwMov(
                    ClsIdCuentaDb_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString)
        Else
            lstrTipoCuenta = "C"
            lstrCuentaCont = ClsPanorama.FobjValorCampo(adrwMov(
                    ClsIdCuentaCr_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString)
        End If
        Dim lblnEsCajaBancos = (lstrCuentaCont = GobjParametros.ObjIdCtaCajaStr.ObjValorPro) OrElse
                GobjParametros.FblnEsCuentaBanco(lstrCuentaCont)
        lstrCuentaCont = lstrCuentaCont.PadRight(10, "0")
        Dim ldecValor As Decimal = adecValor
        ' Fecha
        Dim ldtmFecha As Date = ClsPanorama.FobjValorCampo(adrwMov(ClsFechaNovedadDtm.SstrNombreCampoBd),
                    EnuTipoValor.enuDate)
        Dim lstrFecha = Format(ldtmFecha.Day, "0#") & "/" &
                Format(ldtmFecha.Month, "0#") & "/" &
                ldtmFecha.Year.ToString()
        Dim lstrIdTer As String = If(lblnEsCajaBancos, FstrIdTerceroCajaBancos(adrwMov),
                ClsPanorama.FobjValorCampo(adrwMov(ClsAliasCont_NovStr.SstrNombreCampoBd),
                EnuTipoValor.enuString))
        Dim lstrDetalle = FstrDetalle(adrwMov)
        ' 
        Dim lobjDatos = {lstrCuentaCont, lstrIdTer, astrTipoDoc, lstrPrefDoc, lstrIdDoc,
                "", lstrDetalle, lstrFecha, "", "", "", ldecValor, lstrTipoCuenta, "",
                astrPrefFac, lstrIdFac, ldecBase, "", "", "", "", "", "",
                "", "", "", lstrIdPredAgr, "", "", "", "", "", ""}
        With MstbInterfaz
            .Clear().Append(lstrCuentaCont).Append(CHSTRCOMA).Append(lstrIdTer).Append(CHSTRCOMA)
            .Append(astrTipoDoc).Append(CHSTRCOMA).Append(lstrPrefDoc).Append(CHSTRCOMA)
            .Append(lstrIdDoc).Append(CHSTRCOMA).Append(String.Empty).Append(CHSTRCOMA)
            .Append(lstrDetalle).Append(CHSTRCOMA).Append(lstrFecha).Append(CHSTRCOMA)
            .Append(String.Empty).Append(CHSTRCOMA).Append(String.Empty).Append(CHSTRCOMA)
            .Append(String.Empty).Append(CHSTRCOMA).Append(ldecValor).Append(CHSTRCOMA)
            .Append(lstrTipoCuenta).Append(CHSTRCOMA).Append(String.Empty).Append(CHSTRCOMA)
            .Append(astrPrefFac).Append(CHSTRCOMA).Append(lstrIdFac).Append(CHSTRCOMA)
            .Append(ldecBase).Append(CHSTRCOMA).Append(String.Empty).Append(CHSTRCOMA)
            .Append(String.Empty).Append(CHSTRCOMA).Append(String.Empty).Append(CHSTRCOMA)
            .Append(String.Empty).Append(CHSTRCOMA).Append(String.Empty).Append(CHSTRCOMA)
            .Append(String.Empty).Append(CHSTRCOMA).Append(String.Empty).Append(CHSTRCOMA)
            .Append(String.Empty).Append(CHSTRCOMA).Append(String.Empty).Append(CHSTRCOMA)
            .Append(lstrIdPredAgr).Append(CHSTRCOMA).Append(String.Empty).Append(CHSTRCOMA)
            .Append(String.Empty).Append(CHSTRCOMA).Append(String.Empty).Append(CHSTRCOMA)
            .Append(String.Empty).Append(CHSTRCOMA).Append(String.Empty).Append(CHSTRCOMA)
            .Append(String.Empty)
        End With
        aswInterfaz.WriteLine(MstbInterfaz.ToString())
    End Sub
    ''' <summary>
    ''' Escribe el movimiento CR de un item de factura cuando el tercero es diferente del cliente;
    ''' es decir es un proveedor
    ''' </summary>
    ''' <param name="adrwMov">DataRow con el movimiento</param>
    ''' <param name="astrTipoDoc">Tipo doc para SIIGO</param>
    ''' <param name="adecValor">Valor del Movimiento</param>
    ''' <param name="astrIdTercero">Id del Proveedor</param>
    ''' <remarks></remarks>
    Private Sub SEscribaMovCrFac(adrwMov As DataRow, astrTipoDoc As String,
            astrPrefFac As String, aentIdFac As Integer, adecValor As Decimal,
            astrIdTercero As String, aswInterfaz As StreamWriter)
        Dim lstrCuentaCont As String, lstrIdFac As String
        Dim lstrTipoCuenta As String
        Dim lstrIdPredAgr As String = ClsPanorama.FobjValorCampo(adrwMov(
                ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString)
        Dim lstrPref As String = CType(ClsPanorama.FobjValorCampo(adrwMov(
                        ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString), String)
        If IsNothing(lstrPref) Then lstrPref = String.Empty
        Dim lentIdDoc As Integer = ClsPanorama.FobjValorCampo(adrwMov(
                        ClsIdDocOrigenEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
        Dim lstrIdDoc = "'" & lentIdDoc.ToString.PadLeft(7, "0")
        If aentIdFac = 0 Then
            lstrIdFac = ""
        Else
            lstrIdFac = "'" & aentIdFac.ToString().PadLeft(7, "0")
        End If
        lstrTipoCuenta = "C"
        lstrCuentaCont = ClsPanorama.FobjValorCampo(adrwMov(ClsIdCuentaCr_NovStr.SstrNombreCampoBd),
                EnuTipoValor.enuString)
        lstrCuentaCont = lstrCuentaCont.PadRight(10, "0")
        ' Fecha
        Dim ldtmFecha As Date = ClsPanorama.FobjValorCampo(adrwMov(ClsFechaNovedadDtm.SstrNombreCampoBd),
                    EnuTipoValor.enuDate)
        Dim lstrFecha = Format(ldtmFecha.Day, "0#") & "/" &
                Format(ldtmFecha.Month, "0#") & "/" &
                ldtmFecha.Year.ToString()
        Dim lstrDetalle = FstrDetalle(adrwMov)
        ' 
        Dim lobjDatos = {lstrCuentaCont, astrIdTercero, astrTipoDoc, lstrPref, lstrIdDoc,
                "", lstrDetalle, lstrFecha, "", "", "", adecValor, lstrTipoCuenta, "",
                astrPrefFac, lstrIdFac, "0.00", "", "", "", "", "", "",
                "", "", "", lstrIdPredAgr, "", "", "", "", "", ""}
        With MstbInterfaz
            .Clear().Append(lstrCuentaCont).Append(CHSTRCOMA).Append(astrIdTercero).Append(CHSTRCOMA)
            .Append(astrTipoDoc).Append(CHSTRCOMA).Append(lstrPref).Append(CHSTRCOMA)
            .Append(lstrIdDoc).Append(CHSTRCOMA).Append(String.Empty).Append(CHSTRCOMA)
            .Append(lstrDetalle).Append(CHSTRCOMA).Append(lstrFecha).Append(CHSTRCOMA)
            .Append(String.Empty).Append(CHSTRCOMA).Append(String.Empty).Append(CHSTRCOMA)
            .Append(String.Empty).Append(CHSTRCOMA).Append(adecValor).Append(CHSTRCOMA)
            .Append(lstrTipoCuenta).Append(CHSTRCOMA).Append(String.Empty).Append(CHSTRCOMA)
            .Append(astrPrefFac).Append(CHSTRCOMA).Append(lstrIdFac).Append(CHSTRCOMA)
            .Append("0.00").Append(CHSTRCOMA).Append(String.Empty).Append(CHSTRCOMA)
            .Append(String.Empty).Append(CHSTRCOMA).Append(String.Empty).Append(CHSTRCOMA)
            .Append(String.Empty).Append(CHSTRCOMA).Append(String.Empty).Append(CHSTRCOMA)
            .Append(String.Empty).Append(CHSTRCOMA).Append(String.Empty).Append(CHSTRCOMA)
            .Append(String.Empty).Append(CHSTRCOMA).Append(String.Empty).Append(CHSTRCOMA)
            .Append(lstrIdPredAgr).Append(CHSTRCOMA).Append(String.Empty).Append(CHSTRCOMA)
            .Append(String.Empty).Append(CHSTRCOMA).Append(String.Empty).Append(CHSTRCOMA)
            .Append(String.Empty).Append(CHSTRCOMA).Append(String.Empty).Append(CHSTRCOMA)
            .Append(String.Empty)
        End With
        aswInterfaz.WriteLine(MstbInterfaz.ToString())
    End Sub
#End Region
#End Region
#Region "Generales"
    Friend Overrides ReadOnly Property DtmFechaFinInterfazAnterior() As Date
        Get
            If MdtmFechaFinInterfazAnterior = GCDTMFECHANULA Then
                If Not MblnLeidoUltimoRegistro Then
                    Select Case GobjParametros.ObjIdAppContableByt.ObjValorPro
                        Case EnuAppConta.EnuColon
                            SLeaUltimoRegInteAnterExcel()
                        Case Else
                            Throw New ErrorInesperadoPanLException("Tipo de App Contable no esperado")
                    End Select
                End If
            End If
            Return MdtmFechaFinInterfazAnterior
        End Get
    End Property
    Private Sub SLeaUltimoRegInteAnterExcel()
        If Not MblnLeidoUltimoRegistro Then
            MdtmFechaFinInterfazAnterior = GCDTMFECHAMAXI
            MblnLeidoUltimoRegistro = True
        End If
    End Sub
#End Region
End Class

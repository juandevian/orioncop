Imports System.IO
Imports System.Text
Friend Class ClsInterfazContaSIIGO
    Inherits clsCBInterfazContableOri
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
    Friend Property BlnNube As Boolean = False
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
                    MstrNombreArchivo = "InterfazSIIGO_FM_" & lstrFechaDesde & "_" & lstrFechaHasta
                Else
                    MstrNombreArchivo = "InterfazSIIGO_" & lstrFechaDesde & "_" & lstrFechaHasta
                End If
                HstrArchivoSalidaInterfaz = GstrTrayInterfContable & "\" & MstrNombreArchivo
                If My.Computer.FileSystem.FileExists(HstrArchivoSalidaInterfaz & ".xlsx") Then
                    My.Computer.FileSystem.DeleteFile(HstrArchivoSalidaInterfaz & ".xlsx")
                End If
                SGenereIntSIIGO(astrMens)
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
    Private Sub SGenereIntSIIGO(ByRef astrMens As String)
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
        Dim lstrPref As String, lentIdFra As Integer, lstrNroFac As String,
                lstrNroFacAnt = String.Empty
        Dim lstrTipoDoc = GobjParametros.FstrTipoDoc(aenuIdDocumento)
        Dim lstrFiltro = ClsIdTipoDocOrigenByt.SstrNombreCampoBd & " = " & aenuIdDocumento
        Dim ldrwMvimDoc As DataRow() = adtbMovi.Select(lstrFiltro)
        For Each ldrwNov As DataRow In ldrwMvimDoc
            If aenuIdDocumento = EnuIdDocumentoDef.EnuFacturaVenta Then
                lstrPref = CType(ClsPanorama.FobjValorCampo(ldrwNov(
                    ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString), String)
                If IsNothing(lstrPref) Then lstrPref = String.Empty
                lentIdFra = CType(ClsPanorama.FobjValorCampo(ldrwNov(
                        ClsIdDocOrigenEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger), Integer)
                lstrNroFac = ClsPanorama.FstrNumeroDcto(lstrPref, lentIdFra)
                If lstrNroFacAnt <> lstrNroFac Then
                    SProceseFactura(adtbMovi, lstrPref, lentIdFra, aswInterfaz)
                    lstrNroFacAnt = lstrNroFac
                End If
            Else
                ldecValor = ClsPanorama.FobjValorCampo(ldrwNov(ClsValor_NovDec.SstrNombreCampoBd),
                        EnuTipoValor.enuDecimal)
                If BlnNube Then
                    SEscribaMovDocN(ldrwNov, lstrTipoDoc, ldecValor, True, aswInterfaz)
                    SEscribaMovDocN(ldrwNov, lstrTipoDoc, ldecValor, False, aswInterfaz)
                Else
                    SEscribaMovDoc(ldrwNov, lstrTipoDoc, ldecValor, True, aswInterfaz)
                    SEscribaMovDoc(ldrwNov, lstrTipoDoc, ldecValor, False, aswInterfaz)
                End If
            End If
        Next
    End Sub
    Private Sub SProceseFactura(adtbMovi As DataTable, astrPref As String, aentIdFact As Integer,
            aswInterfaz As StreamWriter)
        Dim lstrFiltro = ClsIdTipoDocOrigenByt.SstrNombreCampoBd & " = " &
                EnuIdDocumentoDef.EnuFacturaVenta & " AND " &
                ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd & " = '" & astrPref & "' AND " &
                ClsIdDocOrigenEnt.SstrNombreCampoBd & " = " & aentIdFact
        Dim ldrwNovsFra As DataRow() = adtbMovi.Select(lstrFiltro), ldrwNov As DataRow
        Dim ldecValorFra = 0D, ldecValor As Decimal, lstrIdCtaDb = String.Empty, i = 0,
                lshrIdItemFac As Short
        Dim lstrIdTerceroCtaCr As String
        Dim lstrTipoDoc = GobjParametros.FstrTipoDoc(EnuIdDocumentoDef.EnuFacturaVenta)
        Dim ldrwNovDbFra As DataRow = Nothing
        Dim lenuTipoNovDbFra As EnuTipoNov
        Dim lobjFactura As New ClsFactura()
        Dim lobjValorLlave As Object()
        Do
            ldrwNov = ldrwNovsFra(i)
            lenuTipoNovDbFra = ClsPanorama.FobjValorCampo(ldrwNov(ClsIdTipoNovedadByt.SstrNombreCampoBd),
                    EnuTipoValor.enuByte)
            If lstrIdCtaDb <> ClsPanorama.FobjValorCampo(ldrwNov(ClsIdCuentaDb_NovStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString) Then
                If Not String.IsNullOrEmpty(lstrIdCtaDb) Then
                    If BlnNube Then
                        SEscribaMovDocN(ldrwNovsFra(i - 1), lstrTipoDoc, ldecValorFra, True,
                                aswInterfaz)
                    Else
                        SEscribaMovDoc(ldrwNovsFra(i - 1), lstrTipoDoc, ldecValorFra, True,
                                aswInterfaz)
                    End If
                End If
                lstrIdCtaDb = ClsPanorama.FobjValorCampo(ldrwNov(ClsIdCuentaDb_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
                ldecValorFra = ClsPanorama.FobjValorCampo(ldrwNov(ClsValor_NovDec.SstrNombreCampoBd),
                            EnuTipoValor.enuDecimal)
            Else
                ldecValorFra += ClsPanorama.FobjValorCampo(ldrwNov(ClsValor_NovDec.SstrNombreCampoBd),
                            EnuTipoValor.enuDecimal)
            End If
            If lenuTipoNovDbFra = EnuTipoNov.EnuDbCap Then
                ldrwNovDbFra = ldrwNov
            End If
            If i = ldrwNovsFra.Length - 1 Then
                If BlnNube Then
                    SEscribaMovDocN(ldrwNovDbFra, lstrTipoDoc, ldecValorFra, True, aswInterfaz)
                Else
                    SEscribaMovDoc(ldrwNovDbFra, lstrTipoDoc, ldecValorFra, True, aswInterfaz)
                End If
            End If
            i += 1
        Loop While i < ldrwNovsFra.Length
        For Each ldrwNovFac As DataRow In ldrwNovsFra
            ldecValor = ClsPanorama.FobjValorCampo(ldrwNovFac(ClsValor_NovDec.SstrNombreCampoBd),
                        EnuTipoValor.enuDecimal)
            lshrIdItemFac = ClsPanorama.FobjValorCampo(ldrwNovFac(ClsIdItemFacturaShr.SstrNombreCampoBd),
                    EnuTipoValor.enuShort)
            lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, astrPref, aentIdFact}
            lobjFactura.SAbra(lobjValorLlave)
            lstrIdTerceroCtaCr = lobjFactura.FstrIdTerceroCtaCr(lshrIdItemFac)
            If String.IsNullOrEmpty(lstrIdTerceroCtaCr) Then
                If BlnNube Then
                    SEscribaMovDocN(ldrwNovFac, lstrTipoDoc, ldecValor, False, aswInterfaz)
                Else
                    SEscribaMovDoc(ldrwNovFac, lstrTipoDoc, ldecValor, False, aswInterfaz)
                End If
            Else
                If BlnNube Then
                    SEscribaMovCrFacN(ldrwNovFac, lstrTipoDoc, ldecValor, lstrIdTerceroCtaCr,
                            aswInterfaz)
                Else
                    SEscribaMovCrFac(ldrwNovFac, lstrTipoDoc, ldecValor, lstrIdTerceroCtaCr,
                            aswInterfaz)
                End If
            End If
        Next
    End Sub
    Private Sub SProceseRecibosCaja(aswInterfaz As StreamWriter)
        Dim lstrTipoDoc = GobjParametros.FstrTipoDoc(EnuIdDocumentoDef.EnuReciboCaja)
        Dim lstrPrefRC = String.Empty, lentIdRC = 0, lstrPrefRCNov As String, lentIdRCNov As Integer
        Dim lstrFiltro As String, lstrIdCuentaDb = String.Empty, lstrIdCuentaDbNov As String
        Dim ldrwMovRC As DataRow()
        Dim ldtbMovRC = FdtbMoviRC()
        For Each ldrwNov As DataRow In ldtbMovRC.Rows
            lstrPrefRCNov = ClsPanorama.FobjValorCampo(ldrwNov(ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            lentIdRCNov = ClsPanorama.FobjValorCampo(ldrwNov(ClsIdDocOrigenEnt.SstrNombreCampoBd),
                    EnuTipoValor.enuInteger)
            lstrIdCuentaDbNov = ClsPanorama.FobjValorCampo(ldrwNov(ClsIdCuentaDb_NovStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            If lstrPrefRC <> lstrPrefRCNov OrElse lentIdRC <> lentIdRCNov OrElse
                    lstrIdCuentaDb <> lstrIdCuentaDbNov Then
                lstrFiltro = ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd & " = '" & lstrPrefRCNov &
                        "' AND " & ClsIdDocOrigenEnt.SstrNombreCampoBd & " = " & lentIdRCNov &
                        " AND " & ClsIdCuentaDb_NovStr.SstrNombreCampoBd & " = '" &
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
        Dim ldecValor = 0D, ldecValorNov As Decimal
        ' Calcular valor debito 
        For Each ldrwMovNov As DataRow In adrwMoviRC
            ldecValorNov = ClsPanorama.FobjValorCampo(ldrwMovNov(ClsValor_NovDec.SstrNombreCampoBd),
                    EnuTipoValor.enuDecimal)
            ldecValor += ldecValorNov
        Next
        If BlnNube Then
            SEscribaMovDocN(adrwMoviRC(0), astrTipoDoc, ldecValor, True, aswInterfaz)
        Else
            SEscribaMovDoc(adrwMoviRC(0), astrTipoDoc, ldecValor, True, aswInterfaz)
        End If
        For Each ldrwMovNov As DataRow In adrwMoviRC
            ldecValorNov = ClsPanorama.FobjValorCampo(ldrwMovNov(ClsValor_NovDec.SstrNombreCampoBd),
                    EnuTipoValor.enuDecimal)
            If BlnNube Then
                SEscribaMovDocN(ldrwMovNov, astrTipoDoc, ldecValorNov, False, aswInterfaz)
            Else
                SEscribaMovDoc(ldrwMovNov, astrTipoDoc, ldecValorNov, False, aswInterfaz)
            End If
        Next
    End Sub
    Private Sub SProceseNotaR(aenuIdDocumento As EnuIdDocumentoDef, aswInterfaz As StreamWriter)
        Dim ldecValor As Decimal
        Dim lstrTipoDoc = GobjParametros.FstrTipoDoc(aenuIdDocumento)
        Dim ldtbMovAntR = FdtbMoviAnt(aenuIdDocumento)
        For Each ldrwNov As DataRow In ldtbMovAntR.Rows
            ldecValor = ClsPanorama.FobjValorCampo(ldrwNov(ClsValor_NovDec.SstrNombreCampoBd),
                    EnuTipoValor.enuDecimal)
            If BlnNube Then
                SEscribaMovDocN(ldrwNov, lstrTipoDoc, ldecValor, True, aswInterfaz)
                SEscribaMovDocN(ldrwNov, lstrTipoDoc, ldecValor, False, aswInterfaz)
            Else
                SEscribaMovDoc(ldrwNov, lstrTipoDoc, ldecValor, True, aswInterfaz)
                SEscribaMovDoc(ldrwNov, lstrTipoDoc, ldecValor, False, aswInterfaz)
            End If
        Next
    End Sub
    Private Sub SProceseNotasAjuste(aswInterfaz As StreamWriter)
        Dim lstrTipoDoc = GobjParametros.FstrTipoDoc(EnuIdDocumentoDef.EnuNotaAjuste)
        Dim ldecValor As Decimal
        Dim ldtbMovNA = FdtbMoviNotaAjuste()
        For Each ldrwNov As DataRow In ldtbMovNA.Rows
            ldecValor = ClsPanorama.FobjValorCampo(ldrwNov(ClsValor_NovDec.SstrNombreCampoBd),
                        EnuTipoValor.enuDecimal)
            If BlnNube Then
                SEscribaMovDocN(ldrwNov, lstrTipoDoc, ldecValor, True, aswInterfaz)
                SEscribaMovDocN(ldrwNov, lstrTipoDoc, ldecValor, False, aswInterfaz)
            Else
                SEscribaMovDoc(ldrwNov, lstrTipoDoc, ldecValor, True, aswInterfaz)
                SEscribaMovDoc(ldrwNov, lstrTipoDoc, ldecValor, False, aswInterfaz)
            End If
        Next
    End Sub
#Region "Archivo Excell"
    Private Sub SEscribaEncabezado(aswInterfaz As StreamWriter)
        With MstbInterfaz
            .Clear()
            If BlnNube Then
                .Append("Tipo de comprobante").Append(CHSTRCOMA).Append("Consecutivo Comprobante")
                .Append(CHSTRCOMA).Append("Fecha de elaboración").Append(CHSTRCOMA)
                .Append("Código cuenta contable").Append(CHSTRCOMA).Append("Identificación Tercero")
                .Append(CHSTRCOMA).Append("Descripción").Append(CHSTRCOMA).Append("Débito")
                .Append(CHSTRCOMA).Append("Crédito")
            Else
                .Append("TIPO DE COMPROBANTE").Append(CHSTRCOMA).Append("CÓDIGO COMPROBANTE")
                .Append(CHSTRCOMA).Append("NÚMERO DE DOCUMENTO").Append(CHSTRCOMA)
                .Append("CUENTA CONTABLE").Append(CHSTRCOMA).Append("DÉBITO O CRÉDITO")
                .Append(CHSTRCOMA).Append("VALOR DE LA SECUENCIA").Append(CHSTRCOMA)
                .Append("AÑO DEL DOCUMENTO").Append(CHSTRCOMA).Append("MES DEL DOCUMENTO")
                .Append(CHSTRCOMA).Append("DÍA DEL DOCUMENTO").Append(CHSTRCOMA)
                .Append("NIT").Append(CHSTRCOMA).Append("DESCRIPCIÓN DE LA SECUENCIA")
            End If
        End With
        aswInterfaz.WriteLine(MstbInterfaz.ToString())
    End Sub
    Private Sub SEscribaMovDoc(adrwMov As DataRow, astrTipoDoc As String,
            adecValor As Decimal, ablnDb As Boolean, aswInterfaz As StreamWriter)
        Dim lstrTipoDocumento = astrTipoDoc.Substring(0, 1)
        Dim lentCodigoDoc As Integer = astrTipoDoc.Substring(1)
        Dim lentIdDocumento = ClsPanorama.FobjValorCampo(adrwMov(ClsIdDocOrigenEnt.SstrNombreCampoBd),
                EnuTipoValor.enuInteger)
        Dim lstrCuentaCont As String
        Dim lstrTipoCuenta As String
        If ablnDb Then
            lstrTipoCuenta = "D"
            lstrCuentaCont = ClsPanorama.FobjValorCampo(adrwMov(ClsIdCuentaDb_NovStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
        Else
            lstrTipoCuenta = "C"
            lstrCuentaCont = ClsPanorama.FobjValorCampo(adrwMov(ClsIdCuentaCr_NovStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
        End If
        Dim lblnEsCajaBancos = (lstrCuentaCont = GobjParametros.ObjIdCtaCajaStr.ObjValorPro) OrElse
                GobjParametros.FblnEsCuentaBanco(lstrCuentaCont)
        lstrCuentaCont = lstrCuentaCont.PadRight(10, "0")
        Dim ldecValor As Decimal = adecValor
        ' Fecha
        Dim ldtmFecha As Date = ClsPanorama.FobjValorCampo(adrwMov(ClsFechaNovedadDtm.SstrNombreCampoBd),
                    EnuTipoValor.enuDate)
        Dim lentAno = ldtmFecha.Year
        Dim lentMes = ldtmFecha.Month
        Dim lentdia = ldtmFecha.Day
        Dim lstrIdTer As String
        If lblnEsCajaBancos Then
            lstrIdTer = FstrIdTerceroCajaBancos(adrwMov)
        Else
            lstrIdTer = ClsPanorama.FobjValorCampo(adrwMov(ClsAliasCont_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
        End If
        Dim lstrDetalle = FstrDetalle(adrwMov)
        With MstbInterfaz
            .Clear.Append(lstrTipoDocumento).Append(CHSTRCOMA).Append(lentCodigoDoc).Append(CHSTRCOMA)
            .Append(lentIdDocumento).Append(CHSTRCOMA).Append(lstrCuentaCont).Append(CHSTRCOMA)
            .Append(lstrTipoCuenta).Append(CHSTRCOMA).Append(ldecValor).Append(CHSTRCOMA)
            .Append(lentAno).Append(CHSTRCOMA).Append(lentMes).Append(CHSTRCOMA)
            .Append(lentdia).Append(CHSTRCOMA).Append(lstrIdTer).Append(CHSTRCOMA)
            .Append(lstrDetalle)
        End With
        aswInterfaz.WriteLine(MstbInterfaz.ToString())
    End Sub
    Private Sub SEscribaMovDocN(adrwMov As DataRow, astrTipoCom As String,
            adecValor As Decimal, ablnDb As Boolean, aswInterfaz As StreamWriter)
        Dim lentTipoCom As Integer = CType(astrTipoCom, Integer)
        Dim lentIdDocumento = ClsPanorama.FobjValorCampo(adrwMov(ClsIdDocOrigenEnt.SstrNombreCampoBd),
                EnuTipoValor.enuInteger)
        Dim lstrCuentaCont As String
        Dim ldecDb = 0D, ldecCr = 0D
        If ablnDb Then
            ldecDb = adecValor
            lstrCuentaCont = ClsPanorama.FobjValorCampo(adrwMov(ClsIdCuentaDb_NovStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
        Else
            ldecCr = adecValor
            lstrCuentaCont = ClsPanorama.FobjValorCampo(adrwMov(ClsIdCuentaCr_NovStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
        End If
        Dim lblnEsCajaBancos = (lstrCuentaCont = GobjParametros.ObjIdCtaCajaStr.ObjValorPro) OrElse
                GobjParametros.FblnEsCuentaBanco(lstrCuentaCont)
        lstrCuentaCont = lstrCuentaCont.PadRight(10, "0")
        ' Fecha
        Dim ldtmFecha As Date = ClsPanorama.FobjValorCampo(adrwMov(ClsFechaNovedadDtm.SstrNombreCampoBd),
                    EnuTipoValor.enuDate)
        Dim lstrIdTer As String
        If lblnEsCajaBancos Then
            lstrIdTer = FstrIdTerceroCajaBancos(adrwMov)
        Else
            lstrIdTer = ClsPanorama.FobjValorCampo(adrwMov(ClsAliasCont_NovStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
        End If
        Dim lstrDetalle = FstrDetalle(adrwMov)
        ' 
        With MstbInterfaz
            .Clear().Append(lentTipoCom).Append(CHSTRCOMA).Append(lentIdDocumento).Append(CHSTRCOMA)
            .Append(ldtmFecha).Append(CHSTRCOMA).Append(lstrCuentaCont).Append(CHSTRCOMA)
            .Append(lstrIdTer).Append(CHSTRCOMA).Append(lstrDetalle).Append(CHSTRCOMA)
            .Append(ldecDb).Append(CHSTRCOMA).Append(ldecCr)
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
                adecValor As Decimal, astrIdTercero As String, aswInterfaz As StreamWriter)
        Dim lstrTipoDocumento = astrTipoDoc.Substring(0, 1)
        Dim lentCodigoDoc As Integer = CType(astrTipoDoc.Substring(1), Integer)
        Dim lentIdDocumento = ClsPanorama.FobjValorCampo(adrwMov(ClsIdDocOrigenEnt.SstrNombreCampoBd),
                EnuTipoValor.enuInteger)
        Dim lstrCuentaCont As String
        Dim lstrTipoCuenta As String
        lstrTipoCuenta = "C"
        lstrCuentaCont = ClsPanorama.FobjValorCampo(adrwMov(ClsIdCuentaCr_NovStr.SstrNombreCampoBd),
                EnuTipoValor.enuString)
        lstrCuentaCont = lstrCuentaCont.PadRight(10, "0")
        ' Fecha
        Dim ldtmFecha As Date = ClsPanorama.FobjValorCampo(adrwMov(ClsFechaNovedadDtm.SstrNombreCampoBd),
                    EnuTipoValor.enuDate)
        Dim lentAno = ldtmFecha.Year
        Dim lentMes = ldtmFecha.Month
        Dim lentdia = ldtmFecha.Day
        Dim lstrDetalle = FstrDetalle(adrwMov)
        ' 
        With MstbInterfaz
            .Clear().Append(lstrTipoDocumento).Append(CHSTRCOMA).Append(lentCodigoDoc)
            .Append(CHSTRCOMA).Append(lentIdDocumento).Append(CHSTRCOMA).Append(lstrCuentaCont)
            .Append(CHSTRCOMA).Append(lstrTipoCuenta).Append(CHSTRCOMA).Append(adecValor)
            .Append(CHSTRCOMA).Append(lentAno).Append(CHSTRCOMA).Append(lentMes).Append(CHSTRCOMA)
            .Append(lentdia).Append(CHSTRCOMA).Append(astrIdTercero).Append(CHSTRCOMA)
            .Append(lstrDetalle)
        End With
        aswInterfaz.WriteLine(MstbInterfaz.ToString())
    End Sub
    ''' <summary>
    ''' Escribe el movimiento CR de un item de factura cuando el tercero es diferente del cliente;
    ''' es decir es un proveedor
    ''' </summary>
    ''' <param name="adrwMov">DataRow con el movimiento</param>
    ''' <param name="astrTipoCom">Tipo doc para SIIGO</param>
    ''' <param name="adecValor">Valor del Movimiento</param>
    ''' <param name="astrIdTercero">Id del Proveedor</param>
    ''' <remarks></remarks>
    Private Sub SEscribaMovCrFacN(adrwMov As DataRow, astrTipoCom As String,
                adecValor As Decimal, astrIdTercero As String, aswInterfaz As StreamWriter)
        Dim lentTipoCom As Integer = CType(astrTipoCom, Integer)
        Dim lentIdDocumento = ClsPanorama.FobjValorCampo(adrwMov(ClsIdDocOrigenEnt.SstrNombreCampoBd),
                EnuTipoValor.enuInteger)
        Dim lstrCuentaCont As String, ldecDb = 0D, ldecCr = adecValor
        lstrCuentaCont = ClsPanorama.FobjValorCampo(adrwMov(ClsIdCuentaCr_NovStr.SstrNombreCampoBd),
                EnuTipoValor.enuString)
        ' Fecha
        Dim ldtmFecha As Date = ClsPanorama.FobjValorCampo(adrwMov(ClsFechaNovedadDtm.SstrNombreCampoBd),
                    EnuTipoValor.enuDate)
        Dim lstrDetalle = FstrDetalle(adrwMov)
        ' 
        With MstbInterfaz
            .Clear().Append(lentTipoCom).Append(CHSTRCOMA).Append(lentIdDocumento).Append(CHSTRCOMA)
            .Append(ldtmFecha).Append(CHSTRCOMA).Append(lstrCuentaCont).Append(CHSTRCOMA)
            .Append(astrIdTercero).Append(CHSTRCOMA).Append(lstrDetalle).Append(CHSTRCOMA)
            .Append(ldecDb).Append(CHSTRCOMA).Append(ldecCr)
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
                        Case EnuAppConta.EnuSIIGO, EnuAppConta.EnuSIIGON
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

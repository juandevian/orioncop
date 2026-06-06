Friend Class ClsImportarOrion
#Region "Definiciones"
#Region "Enumeradores"
    Private Enum EnuTipoObjetoDef As Integer
        None = 0
        enuAgrupador
        enuServicio
    End Enum
#End Region
    Public Event EvnInicioImportacion As EventHandler(Of ClsPanEventArgs)
    Public Event EvnAvance As EventHandler(Of ClsPanEventArgs)
    Private MobjArgumentoEventoPan As ClsPanEventArgs = Nothing
    Private ReadOnly Mobjcliente As New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
#End Region
#Region "Importar facturas iniciales desde Excel"
    Friend ReadOnly Property ObjArgumentoEventoPan As ClsPanEventArgs
        Get
            If IsNothing(mobjArgumentoEventoPan) Then
                MobjArgumentoEventoPan = New ClsPanEventArgs With {
                    .BlnCancele = False,
                    .BlnVaciandoObjeto = False
                }
            End If
            Return mobjArgumentoEventoPan
        End Get
    End Property
    Friend Function FblnImportoFacturasIniciales(ByRef astrMens As String) As Boolean
        Dim lblnNoHayError = False, lblnImportoFac = False
        Try
            ClsOrionCop.BlnProcesoEspecial = True
            GobjPanDat.SControleProcesoObj(True)
            GobjPanDat.SInicialiceTransaccion()
            lblnImportoFac = FblnImportoFacturasInici(astrMens)
            lblnNoHayError = True
        Catch ex As PanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        Finally
            ClsOrionCop.BlnProcesoEspecial = False
            If lblnNoHayError AndAlso lblnImportoFac Then
                GobjPanDat.SConfirmeTransaccion()
                GobjPanDat.SControleProcesoObj(False)
            Else
                lblnImportoFac = False
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
        Return lblnImportoFac
    End Function
    Private Function FblnImportoFacturasInici(ByRef astrMens As String) As Boolean
        Dim lstrArchivo = GstrTrayDatPrg & "PlantillaFacturas_OrionPLus.xlsx"
        If Not My.Computer.FileSystem.FileExists(lstrArchivo) Then
            astrMens = "El Archivo de Excel con las Facturas para importar no existe en la " &
                    "Ubicación esperada!"
            Return False
        End If
        Dim ldblCantAProcesar = 0.0, lblnImporto = False
        Dim ldtbItemsFactura = ClsPanorama.FdtbTablaAccess(lstrArchivo, "", "PlantillaFacturas_OrionPLus$")
        If FblnEstaArchImpFactOK(ldtbItemsFactura, ldblCantAProcesar, astrMens) Then
            ObjArgumentoEventoPan.BlnCancele = False
            ObjArgumentoEventoPan.EnuProceso = EnuProcesoDef.enuImpoFras
            ObjArgumentoEventoPan.DblCantAProcesar = ldblCantAProcesar
            RaiseEvent EvnInicioImportacion(Me, ObjArgumentoEventoPan)
            Dim ldrwItesmFacturas() As DataRow = ldtbItemsFactura.Select()
            Dim ldtmFechaCausoMora = GCDTMFECHANULA, ldtmFechaFact = GCDTMFECHANULA
            Dim ldtmFechaPagoSinMora As Date = GCDTMFECHANULA
            Dim ldtmFechaVence = GCDTMFECHANULA, lshrIdAgruFact = 0S, i = 0, j = 0
            Dim lstrPeriodo = String.Empty, lstrPrefFact = String.Empty
            Dim lentIdFact As Integer = 0, lstrIdPredioFac = String.Empty, ldblIdCliente = 0.0
            Dim ldecValorItem = 0D, lshrIdServicio = 0S, lshrIdAno = 0S, lstrIdPredioItem = String.Empty
            Dim lentFormaPago = 0, lentMedioPago = 0
            Dim lblnEsExcluidoIva = False
            Dim ldblTarifaIva = 0.0, lstrDetalleItem = String.Empty
            Dim ldrwItemFact As DataRow = Nothing
            Dim lobjFactura As ClsFactura = Nothing, lobjItemFact As ClsItemFactura = Nothing
            ldrwItemFact = ldrwItesmFacturas(j)
            Do While True
                lentIdFact = ClsPanorama.FobjValorCampo(ldrwItemFact(ClsIdFacturaEnt.SstrNombreCampoBd),
                        EnuTipoValor.enuInteger)
                ObjArgumentoEventoPan.BlnCancele = False
                ObjArgumentoEventoPan.DblCantProcesada = i
                RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
                If lentIdFact = 0 Then
                    j += 1
                    ldrwItemFact = ldrwItesmFacturas(j)
                Else
                    ldblIdCliente = ClsPanorama.FobjValorCampo(ldrwItemFact(ClsIdCliente_FactDbl.SstrNombreCampoBd),
                                EnuTipoValor.enuDouble)
                    Mobjcliente.SAbra({GshrIdCarpeta, GshrIdCentroUtil, ldblIdCliente})
                    lstrPeriodo = ClsPanorama.FobjValorCampo(ldrwItemFact(ClsPeriodo_ItemFactStr.SstrNombreCampoBd),
                            EnuTipoValor.enuString)
                    lstrPrefFact = ClsPanorama.FobjValorCampo(ldrwItemFact(ClsPrefijo_FactStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
                    ldtmFechaFact = ClsPanorama.FobjValorCampo(ldrwItemFact(ClsFechaFacturaDtm.SstrNombreCampoBd),
                            EnuTipoValor.enuDate)
                    ldtmFechaVence = ClsPanorama.FobjValorCampo(ldrwItemFact(ClsFechaVencimientoDtm.SstrNombreCampoBd),
                            EnuTipoValor.enuDate)
                    ldtmFechaPagoSinMora = ClsPanorama.FobjValorCampo(ldrwItemFact(ClsFechaGraciaDtm.SstrNombreCampoBd),
                            EnuTipoValor.enuDate)
                    lstrIdPredioFac = ClsPanorama.FobjValorCampo(ldrwItemFact(ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd),
                            EnuTipoValor.enuString)
                    lentFormaPago = ClsPanorama.FobjValorCampo(ldrwItemFact(ClsIdFormaPagoByt.SstrNombreCampoBd),
                            EnuTipoValor.enuInteger)
                    lentMedioPago = ClsPanorama.FobjValorCampo(ldrwItemFact(ClsIdMedioPagoByt.SstrNombreCampoBd),
                            EnuTipoValor.enuInteger)
                    lobjFactura = Mobjcliente.ObjNuevaFactura(EnuModoFacturacionDef.EnuImportada)
                    With lobjFactura
                        .ObjFechaFacturaDtm.ObjValorPro = ldtmFechaFact
                        .ObjFechaGraciaDtm.ObjValorPro = ldtmFechaPagoSinMora
                        .ObjFechaVencimientoDtm.ObjValorPro = ldtmFechaVence
                        .ObjIdFacturaEnt.ObjValorPro = lentIdFact
                        .ObjIdUsuario_FactStr.ObjValorPro = GstrIdUsuario
                        .ObjPieFacturaDos_FactStr.ObjValorPro =
                                GobjParametros.ObjPieFacturaDosStr.ToString()
                        .ObjPieFacturaUno_FactStr.ObjValorPro =
                                GobjParametros.ObjPieFacturaUnoStr.ToString()
                        .ObjPrefijo_FactStr.ObjValorPro = lstrPrefFact
                        .ObjOrigenInstanciaStr.ObjValorPro = GstrOrigenActual
                        .ObjIdPredioAgrupador_FacStr.ObjValorPro = lstrIdPredioFac
                        .ObjIdFormaPagoByt.ObjValorPro = lentFormaPago
                        .ObjIdMedioPagoByt.ObjValorPro = lentMedioPago
                    End With
                    Do While ClsPanorama.FobjValorCampo(ldrwItemFact(
                            ClsPrefijo_FactStr.SstrNombreCampoBd),
                            EnuTipoValor.enuString) = lstrPrefFact AndAlso
                            ClsPanorama.FobjValorCampo(
                            ldrwItemFact(ClsIdFacturaEnt.SstrNombreCampoBd),
                            EnuTipoValor.enuInteger) = lentIdFact
                        lobjItemFact = lobjFactura.FobjNuevoItemFactura
                        ldecValorItem = ClsPanorama.FobjValorCampo(ldrwItemFact(ClsValor_ItemFactDec.SstrNombreCampoBd),
                                EnuTipoValor.enuDecimal)
                        lshrIdAno = ClsPanorama.FobjValorCampo(ldrwItemFact(ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd),
                                EnuTipoValor.enuShort)
                        lshrIdServicio = ClsPanorama.FobjValorCampo(ldrwItemFact(ClsIdServicio_ItemFactShr.SstrNombreCampoBd),
                                EnuTipoValor.enuShort)
                        lstrIdPredioItem = ClsPanorama.FobjValorCampo(ldrwItemFact(ClsIdPredio_ItemFactStr.SstrNombreCampoBd),
                                EnuTipoValor.enuString)
                        lstrDetalleItem = ClsPanorama.FobjValorCampo(ldrwItemFact(ClsDetalle_ItemFactStr.SstrNombreCampoBd),
                                EnuTipoValor.enuString)
                        lblnEsExcluidoIva = ClsPanorama.FobjValorCampo(ldrwItemFact(ClsEsExcluidoIva_ItemFactBln.SstrNombreCampoBd),
                                EnuTipoValor.enuBoolean)
                        ldblTarifaIva = ClsPanorama.FobjValorCampo(ldrwItemFact(ClsTarifaIva_ItemFactDbl.SstrNombreCampoBd),
                                EnuTipoValor.enuDouble)
                        ldtmFechaCausoMora = ClsPanorama.FobjValorCampo(ldrwItemFact(ClsFechaCausoIntMora_Dtm.SstrNombreCampoBd),
                            EnuTipoValor.enuDate)
                        With lobjItemFact
                            .ObjPeriodo_ItemFactStr.ObjValorPro = lstrPeriodo
                            .ObjDebitos_ItemFactDec.ObjValorPro = ldecValorItem
                            .ObjDetalle_ItemFactStr.ObjValorPro = lstrDetalleItem
                            .ObjEsExcluidoIva_ItemFactBln.ObjValorPro = lblnEsExcluidoIva
                            .ObjEsPrefactura_ItemFactBln.ObjValorPro = False
                            .ObjIdAno_ServicioItemFactShr.ObjValorPro = lshrIdAno
                            .ObjIdPredio_ItemFactStr.ObjValorPro = lstrIdPredioItem
                            .ObjIdServicio_ItemFactShr.ObjValorPro = lshrIdServicio
                            .ObjTarifaIva_ItemFactDbl.ObjValorPro = ldblTarifaIva
                            .ObjValor_ItemFactDec.ObjValorPro = ldecValorItem
                            .ObjFechaGraciaIFDtm.ObjValorPro = ldtmFechaPagoSinMora
                            .ObjFechaVencimientoIFDtm.ObjValorPro = ldtmFechaVence
                            .ObjFechaCausoIntMora_Dtm.ObjValorPro = ldtmFechaCausoMora
                        End With
                        lobjFactura.SAdicioneNuevoItem(lobjItemFact)
                        j += 1
                        i += 1
                        If i = ldblCantAProcesar Then
                            Exit Do
                        Else
                            ldrwItemFact = ldrwItesmFacturas(j)
                        End If
                    Loop
                    lobjFactura.SActualice(True)
                    If i = ldblCantAProcesar Then
                        lblnImporto = True
                        Exit Do
                    End If
                End If
            Loop
        Else
            lblnImporto = False
        End If
        Return lblnImporto
    End Function
    Private Shared Function FblnEstaArchImpFactOK(adtbFacturasIni As DataTable,
                ByRef adblCantidadAFacturar As Double, ByRef astrMens As String) As Boolean
        Dim lblnEstaOk = FblnNoEstaVacioArchImpFact(adtbFacturasIni, adblCantidadAFacturar)
        If Not lblnEstaOk Then
            astrMens = "El Archivo de Excel esta vacio!"
            Return lblnEstaOk
        End If
        lblnEstaOk = FblnExistenLosServicios(adtbFacturasIni, astrMens)
        If Not lblnEstaOk Then
            Return lblnEstaOk
        End If
        lblnEstaOk = FblnClienteYPredioExisten(adtbFacturasIni, astrMens)
        If Not lblnEstaOk Then
            Return lblnEstaOk
        End If
        lblnEstaOk = FblnFechasOK(adtbFacturasIni, astrMens)
        Return lblnEstaOk
    End Function
    Private Shared Function FblnNoEstaVacioArchImpFact(adtbFacturasIni As DataTable,
            ByRef adblCantidadAProcesar As Double) As Boolean
        Dim lblnNoEstaVacio = (adtbFacturasIni.Rows.Count > 0)
        Dim ldblCanFact As Double, lentIdFact As Integer
        If lblnNoEstaVacio Then
            For Each ldrwItemFact As DataRow In adtbFacturasIni.Rows
                lentIdFact = ClsPanorama.FobjValorCampo(ldrwItemFact(ClsIdFacturaEnt.SstrNombreCampoBd),
                        EnuTipoValor.enuInteger)
                If lentIdFact > 0 Then
                    ldblCanFact += 1
                End If
            Next
        End If
        If ldblCanFact > 0 Then
            adblCantidadAProcesar = ldblCanFact
        Else
            lblnNoEstaVacio = False
        End If
        Return lblnNoEstaVacio
    End Function
    Private Shared Function FblnClienteYPredioExisten(adtbFacturasImportar As DataTable,
            ByRef astrMens As String) As Boolean
        Dim ldblIdCliente As Double, lstrIdPredioFac As String, lstrIdPredioItem As String
        Dim lstrDetalle As String, lentIdFact As Integer
        Dim lobjCliente As New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
        Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
        Dim lblnEstaOk = False
        For Each ldrwItemFac As DataRow In adtbFacturasImportar.Rows
            lentIdFact = ClsPanorama.FobjValorCampo(ldrwItemFac(ClsIdFacturaEnt.SstrNombreCampoBd),
                        EnuTipoValor.enuInteger)
            If lentIdFact > 0 Then
                ldblIdCliente = ClsPanorama.FobjValorCampo(ldrwItemFac(ClsIdCliente_FactDbl.SstrNombreCampoBd),
                           EnuTipoValor.enuDouble)
                lstrIdPredioFac = ClsPanorama.FobjValorCampo(ldrwItemFac(ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
                lstrIdPredioItem = ClsPanorama.FobjValorCampo(ldrwItemFac(ClsIdPredio_ItemFactStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
                lstrDetalle = ClsPanorama.FobjValorCampo(ldrwItemFac(ClsDetalle_ItemFactStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
                lblnEstaOk = Not (String.IsNullOrEmpty(lstrDetalle) OrElse
                        String.IsNullOrEmpty(lstrDetalle))
                If Not lblnEstaOk Then
                    astrMens = "El detalle de la Factura Nro. " & lentIdFact.ToString & " no es válido!"
                    Exit For
                End If
                lobjCliente.SAbra({GshrIdCarpeta, GshrIdCentroUtil, ldblIdCliente})
                lblnEstaOk = lobjCliente.BlnExiste
                If Not lblnEstaOk Then
                    astrMens = "El Cliente de la Factura Nro. " & lentIdFact.ToString & " no está creado!"
                    Exit For
                End If
                If Not String.IsNullOrEmpty(lstrIdPredioFac) Then
                    lobjPredio.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrIdPredioFac})
                    lblnEstaOk = lobjPredio.BlnExiste
                    If Not lblnEstaOk Then
                        astrMens = "El Predio Agrupador de la Factura Nro. " & lentIdFact.ToString +
                                " no está creado!"
                        Exit For
                    End If
                    lobjPredio.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrIdPredioItem})
                    lblnEstaOk = lobjPredio.BlnExiste
                    If Not lblnEstaOk Then
                        astrMens = "El Predio del Item de la Factura Nro. " & lentIdFact.ToString +
                                " no está creado!"
                        Exit For
                    Else
                        If lobjPredio.ObjIdPredioAgrupadorStr.ObjValorPro <> lstrIdPredioFac Then
                            astrMens = "El Predio Agrupador de la Factura Nro. " & lentIdFact.ToString +
                                    " no es el mismo del Predio Agrupador del Predio del Item!"
                            Exit For
                        End If
                    End If
                End If
            End If
        Next
        Return lblnEstaOk
    End Function
    Private Shared Function FblnExistenLosServicios(adtbFacturasImportar As DataTable,
            ByRef astrMens As String) As Boolean
        Dim lshrIdAno As Short, lshrIdServicio As Short, lentIdFact As Integer
        Dim lblnExiste = True
        Dim lobjServicio As ClsServicio
        For Each ldrwItemFact As DataRow In adtbFacturasImportar.Rows
            lentIdFact = ClsPanorama.FobjValorCampo(ldrwItemFact(ClsIdFacturaEnt.SstrNombreCampoBd),
                EnuTipoValor.enuInteger)
            If lentIdFact > 0 Then
                lshrIdAno = ClsPanorama.FobjValorCampo(ldrwItemFact(ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd),
                        EnuTipoValor.enuShort)
                lshrIdServicio = ClsPanorama.FobjValorCampo(ldrwItemFact(ClsIdServicio_ItemFactShr.SstrNombreCampoBd),
                        EnuTipoValor.enuShort)
                Dim lstrKey = lshrIdAno.ToString & "," & lshrIdServicio.ToString
                If lshrIdAno > 0 Then
                    If lshrIdAno <> GobjParametros.ObjAnoActual.ObjIdAnoShr.ObjValorPro Then
                        astrMens = "El Año correspondiente a la Factura Nro. " &
                                lentIdFact.ToString & " no es el Año actual!"
                        lblnExiste = False
                        Exit For
                    End If
                    If Not GobjParametros.ObjAnoActual.ColServiciosAno.Contains(lstrKey) Then
                        lblnExiste = False
                        astrMens = "El Servicio correspondiente a la Factura Nro. " &
                                lentIdFact.ToString & " no existe!"
                        Exit For
                    Else
                        lobjServicio = GobjParametros.ObjAnoActual.ColServiciosAno(lstrKey)
                    End If
                Else
                    If Not GobjParametros.ColServiciosPer.Contains(lstrKey) Then
                        lblnExiste = False
                        astrMens = "El Servicio correspondiente a la Factura Nro. " &
                                lentIdFact.ToString & " no existe!"
                        Exit For
                    Else
                        lobjServicio = GobjParametros.ColServiciosPer(lstrKey)
                    End If
                End If
            End If
        Next
        Return lblnExiste
    End Function
    Private Shared Function FblnFechasOK(adtbFacturasImportar As DataTable,
            ByRef astrMens As String) As Boolean
        Dim lblnFechasOk = True, lentIdFact As Integer
        Dim lstrPeriodo As String, ldecValorItem As Decimal
        Dim ldtmFechaFact As Date, ldtmFechaVence As Date, lstrPrefFac As String
        Dim lstrPerActual = GobjParametros.ObjAnoActual.StrIdPeriodoActual
        Dim ldtmfechaMaxPago As Date
        For Each ldrwItemFac As DataRow In adtbFacturasImportar.Rows
            lentIdFact = ClsPanorama.FobjValorCampo(ldrwItemFac(ClsIdFacturaEnt.SstrNombreCampoBd),
                        EnuTipoValor.enuInteger)
            If lentIdFact > 0 Then
                lstrPeriodo = ClsPanorama.FobjValorCampo(ldrwItemFac(ClsPeriodo_ItemFactStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
                ldtmFechaFact = ClsPanorama.FobjValorCampo(ldrwItemFac(ClsFechaFacturaDtm.SstrNombreCampoBd),
                        EnuTipoValor.enuDate)
                ldtmFechaVence = ClsPanorama.FobjValorCampo(ldrwItemFac(ClsFechaVencimientoDtm.SstrNombreCampoBd),
                        EnuTipoValor.enuDate)
                ldtmfechaMaxPago = ClsPanorama.FobjValorCampo(ldrwItemFac(ClsFechaGraciaDtm.SstrNombreCampoBd),
                        EnuTipoValor.enuDate)
                If ldtmFechaFact.Year.ToString & Format(ldtmFechaFact.Month, "0#") <>
                        lstrPerActual Then
                    lblnFechasOk = False
                    astrMens = "La fecha de la factura Nro. " & lentIdFact.ToString & " debe " &
                            "corresponder al Período Actual!"
                    Exit For
                End If
                If ldtmFechaVence < ldtmFechaFact Then
                        lblnFechasOk = False
                        astrMens = "La Fecha de vencimiento de la Factura Nro. " &
                            lentIdFact.ToString & " es anterior a la Fecha de la Factura!"
                        Exit For
                    End If
                    If ClsOrionCop.FstrPeriodoDeFecha(ldtmFechaFact) <> lstrPeriodo Then
                        lblnFechasOk = False
                        astrMens = "El Período y la Fecha de la Factura Nro. " &
                            lentIdFact.ToString & " no se corresponden!"
                        Exit For
                    End If
                    If ldtmfechaMaxPago < ldtmFechaVence Then
                        lblnFechasOk = False
                        astrMens = "La Fecha de pago sin Mora es anterior a la fecha de Vencimiento " &
                            "en la Factura Nro. " & lentIdFact.ToString & "!"
                        Exit For
                    End If
                    ldecValorItem = ClsPanorama.FobjValorCampo(ldrwItemFac(
                        ClsValor_ItemFactDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
                    lstrPrefFac = ClsPanorama.FobjValorCampo(ldrwItemFac(ClsPrefijo_FactStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
                    If ldecValorItem - Int(ldecValorItem) > 0 Then
                        lblnFechasOk = False
                        Dim lstrNroFac = ClsPanorama.FstrNumeroDcto(lstrPrefFac, lentIdFact)
                        astrMens = "El valor de la Factura " & lstrNroFac & " tiene Centavos!"
                        Exit For
                    End If
                End If
        Next
        Return lblnFechasOk
    End Function
#End Region
#Region "Varios"
    Friend Shared Function FblnTienePermisos() As Boolean
        Dim lblnTienePermisos As Boolean = True
        With gobjPanorama.objUsuarioActual
            lblnTienePermisos = .fblnTienePermiso(EnuIdClasesPanDef.enuAno, enuIdAccionDef.enuCrear)
            If lblnTienePermisos Then
                lblnTienePermisos = .fblnTienePermiso(EnuIdClasesPanDef.enuAgrServicios, enuIdAccionDef.enuCrear)
            End If
            If lblnTienePermisos Then
                lblnTienePermisos = .fblnTienePermiso(EnuIdClasesPanDef.enuServicio, enuIdAccionDef.enuCrear)
            End If
            If lblnTienePermisos Then
                lblnTienePermisos = .fblnTienePermiso(EnuIdClasesPanDef.enuFactura, enuIdAccionDef.enuCrear)
            End If
            If lblnTienePermisos Then
                lblnTienePermisos = .fblnTienePermiso(EnuIdClasesPanDef.enuAnticipo, enuIdAccionDef.enuCrear)
            End If
        End With
        Return lblnTienePermisos
    End Function
#End Region
End Class

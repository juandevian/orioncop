Friend Class ClsCentroUtilOriCop
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriCentrosUtilidadOriCop"
    '
    Private McolSectores As Collection = Nothing
    Private McolAnos As Collection = Nothing
    Private MobjAnoActual As ClsAno = Nothing
    Private McolModulos As Collection = Nothing
    Private McolAgrupadoresServicios As Collection = Nothing
    Private MdtbAgrupadoresServicios As DataTable = Nothing
    Private McolServiciosPer As Collection = Nothing
    Private MdtbServiciosPer As DataTable = Nothing
    Private McolFechasServicio As Collection = Nothing
    Private ReadOnly ObjServicioPerNuevo As ClsServicio = Nothing
    ' Documentos Contables
    Private McolDocumentos As Collection = Nothing
    Private MobjDocumento As ClsDocumento = Nothing
    Private MobjProveedorEFac As ClsProveedorEFac = Nothing
    Private MenuEstadoInstalacion As EnuEstadoInstalacion = EnuEstadoInstalacion.None
#End Region

#Region "Constructores"
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="adrwObjeto">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Friend Sub New(adrwObjeto As DataRow)
        HobjPadre = Nothing
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsAnulable = False
        HblnEsSuprimible = False
        '
        DrwRegistroActual = adrwObjeto
        DtbTablaColeccion = DrwRegistroActual.Table
    End Sub
#End Region

#Region "Propiedades"
#Region "Propiedades indentificadoras"
    Protected Overrides ReadOnly Property HstrNombreTabla As String
        Get
            Return MCSTRNOMBRETABLA
        End Get
    End Property
    Protected Overrides ReadOnly Property HenuIdClase As EnuIdClasesPanDef
        Get
            Return EnuIdClasesPanDef.EnuCenutiliOriCop
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Copropiedad Orión Plus"
        End Get
    End Property
    Friend Shared ReadOnly Property SstrNombreTabla As String
        Get
            Return MCSTRNOMBRETABLA
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjAutorizaEFacBln As New ClsAutorizaEFacBln(Me)
    Friend ReadOnly Property ObjIdAppContableByt As New ClsIdAppContableByt(Me)
    Friend ReadOnly Property ObjBaseRedondeoCPByt As New ClsBaseRedondeoCPByt(Me)
    Friend ReadOnly Property ObjBaseRedondeoGeneralDbl As New ClsBaseRedondeoGeneralDbl(Me)
    Friend ReadOnly Property ObjBaseRedondeoIntMoraDbl As New ClsBaseRedondeoIntMoraDbl(Me)
    Friend ReadOnly Property ObjCodigoEmpShr As New ClsCodigoEmpShr(Me)
    Friend ReadOnly Property ObjConsolidaItemsFacBln As New ClsConsolidaItemsFacBln(Me)
    Friend ReadOnly Property ObjDiasParaPersuasivoShr As New ClsDiasParaPersuasivoShr(Me)
    Friend ReadOnly Property ObjDiasParaPrejuridicoShr As New ClsDiasParaPrejuridicoShr(Me)
    Friend ReadOnly Property ObjDiasParaJuridicoShr As New ClsDiasParaJuridicoShr(Me)
    Friend ReadOnly Property ObjDiasParaPerdidaShr As New ClsDiasParaPerdidaShr(Me)
    Friend ReadOnly Property ObjExigeFechaHoyCajaBln As New ClsExigeFechaHoyCajaBln(Me)
    Friend ReadOnly Property ObjExigeFechaHoyDocsBln As New ClsExigeFechaHoyDocsBln(Me)
    Friend ReadOnly Property ObjFechaResolucionContDtm As New ClsFechaResolucionContDtm(Me)
    Friend ReadOnly Property ObjFechaResolucionFactDtm As New ClsFechaResolucionFactDtm(Me)
    Friend ReadOnly Property ObjFechaUltCausacionGralDtm As New ClsFechaUltCausacionGralDtm(Me)
    Friend ReadOnly Property ObjFechaVenceResolFactDtm As New ClsFechaVenceResolFactDtm(Me)
    Friend ReadOnly Property ObjFirmaRCeMail As New ClsFirmaRCeMail(Me)
    Friend ReadOnly Property ObjIdCarpetaCentroUtilOriCopShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_OriCopShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdCtaAnticiposRecibidosStr As New ClsIdCtaAnticiposRecibidosStr(Me)
    Friend ReadOnly Property ObjIdCtaCajaStr As New ClsIdCtaCajaStr(Me)
    Friend ReadOnly Property ObjIdCtaDescuentosPPStr As New ClsIdCtaDescuentosPPStr(Me)
    Friend ReadOnly Property ObjIdCtaImptosAsumidosStr As New ClsIdCtaImptosAsumidosStr(Me)
    Friend ReadOnly Property ObjIdCtaIngPorIdentificarStr As New ClsIdCtaIngPorIdentificarStr(Me)
    Friend ReadOnly Property ObjIdCtaIntMoraDbStr As New ClsIdCtaIntMoraDbStr(Me)
    Friend ReadOnly Property ObjIdCtaReteFuenteStr As New ClsIdCtaReteFuenteStr(Me)
    Friend ReadOnly Property ObjIdCtaReteIcaStr As New ClsIdCtaReteIcaStr(Me)
    Friend ReadOnly Property ObjIdCtaReteIvaStr As New ClsIdCtaReteIvaStr(Me)
    Friend ReadOnly Property ObjIdProvEFacByt As New ClsIdProvEFacByt(Me)
    Friend ReadOnly Property ObjInformaSaldoTotalDespuesRCBln As New ClsInformaSaldoTotalDespuesRCBln(Me)
    Friend ReadOnly Property ObjNoMostrarAyudaBln As New ClsNoMostrarAyudaBln(Me)
    Friend ReadOnly Property ObjNotificacionesSonorasBln As New ClsNotificacionesSonorasBln(Me)
    Friend ReadOnly Property ObjNumeroResolContiStr As New ClsNumeroResolContiStr(Me)
    Friend ReadOnly Property ObjNumeroResolFacturaStr As New ClsNumeroResolFacturaStr(Me)
    Friend ReadOnly Property ObjParametrizacionOkBln As New ClsParametrizacionOkBln(Me)
    Friend ReadOnly Property ObjPermiteAnticipoPorServicioBln As New ClsPermiteAnticipoPorServicioBln(Me)
    Friend ReadOnly Property ObjPieFacturaDosStr As New ClsPieFacturaDos_CUStr(Me)
    Friend ReadOnly Property ObjPieFacturaUnoStr As New ClsPieFacturaUno_CUStr(Me)
    Friend ReadOnly Property ObjPlazoDefectoFacManualShr As New ClsPlazoDefectoFacManualShr(Me)
    Friend ReadOnly Property ObjPrefijoFactContStr As New ClsPrefijoFactContStr(Me)
    Friend ReadOnly Property ObjRangoFraConFinEnt As New ClsRangoFraConFinEnt(Me)
    Friend ReadOnly Property ObjRangoFraConIniEnt As New ClsRangoFraConIniEnt(Me)
    Friend ReadOnly Property ObjRangoFraFinEnt As New ClsRangoFraFinEnt(Me)
    Friend ReadOnly Property ObjRangoFraIniEnt As New ClsRangoFraIniEnt(Me)
    Friend ReadOnly Property ObjServicioIdActivoBln As New ClsServicioIdActivoBln(Me)
    Friend ReadOnly Property ObjTarifaReteIvaDbl As New ClsTarifaReteIvaDbl(Me)
    Friend ReadOnly Property ObjTipoInterfazByt As New ClsTipoInterfazByt(Me)
    Friend ReadOnly Property ObjIdMedioPagoDefectoByt As New ClsIdMedioPagoDefectoByt(Me)
    Friend ReadOnly Property ObjTipoTerceroCajaByt As New ClsTipoTerceroCajaByt(Me)
    Friend ReadOnly Property ObjTotalAreaCopropDec As New ClsTotalAreaCopropDec(Me)
    Friend ReadOnly Property ObjTotalAreaPondDec As New ClsTotalAreaPondDec(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjAutorizaEFacBln)
                HcolPropiedades.Add(ObjBaseRedondeoCPByt)
                HcolPropiedades.Add(ObjBaseRedondeoGeneralDbl)
                HcolPropiedades.Add(ObjBaseRedondeoIntMoraDbl)
                HcolPropiedades.Add(ObjCodigoEmpShr)
                HcolPropiedades.Add(ObjConsolidaItemsFacBln)
                HcolPropiedades.Add(ObjDiasParaPersuasivoShr)
                HcolPropiedades.Add(ObjDiasParaPrejuridicoShr)
                HcolPropiedades.Add(ObjDiasParaJuridicoShr)
                HcolPropiedades.Add(ObjDiasParaPerdidaShr)
                HcolPropiedades.Add(ObjExigeFechaHoyCajaBln)
                HcolPropiedades.Add(ObjExigeFechaHoyDocsBln)
                HcolPropiedades.Add(ObjFechaResolucionContDtm)
                HcolPropiedades.Add(ObjFechaResolucionFactDtm)
                HcolPropiedades.Add(ObjFechaUltCausacionGralDtm)
                HcolPropiedades.Add(ObjFechaVenceResolFactDtm)
                HcolPropiedades.Add(ObjFirmaRCeMail)
                HcolPropiedades.Add(ObjIdAppContableByt)
                HcolPropiedades.Add(ObjIdCarpetaCentroUtilOriCopShr)
                HcolPropiedades.Add(ObjIdCentroUtil_OriCopShr)
                HcolPropiedades.Add(ObjIdCtaAnticiposRecibidosStr)
                HcolPropiedades.Add(ObjIdCtaCajaStr)
                HcolPropiedades.Add(ObjIdCtaDescuentosPPStr)
                HcolPropiedades.Add(ObjIdCtaImptosAsumidosStr)
                HcolPropiedades.Add(ObjIdCtaIngPorIdentificarStr)
                HcolPropiedades.Add(ObjIdCtaIntMoraDbStr)
                HcolPropiedades.Add(ObjIdCtaReteFuenteStr)
                HcolPropiedades.Add(ObjIdCtaReteIcaStr)
                HcolPropiedades.Add(ObjIdCtaReteIvaStr)
                HcolPropiedades.Add(ObjIdProvEFacByt)
                HcolPropiedades.Add(ObjInformaSaldoTotalDespuesRCBln)
                HcolPropiedades.Add(ObjNotificacionesSonorasBln)
                HcolPropiedades.Add(ObjNoMostrarAyudaBln)
                HcolPropiedades.Add(ObjNumeroResolContiStr)
                HcolPropiedades.Add(ObjNumeroResolFacturaStr)
                HcolPropiedades.Add(ObjParametrizacionOkBln)
                HcolPropiedades.Add(ObjPermiteAnticipoPorServicioBln)
                HcolPropiedades.Add(ObjPrefijoFactContStr)
                HcolPropiedades.Add(ObjPieFacturaDosStr)
                HcolPropiedades.Add(ObjPieFacturaUnoStr)
                HcolPropiedades.Add(ObjPlazoDefectoFacManualShr)
                HcolPropiedades.Add(ObjRangoFraConFinEnt)
                HcolPropiedades.Add(ObjRangoFraConIniEnt)
                HcolPropiedades.Add(ObjRangoFraFinEnt)
                HcolPropiedades.Add(ObjRangoFraIniEnt)
                HcolPropiedades.Add(ObjServicioIdActivoBln)
                HcolPropiedades.Add(ObjTarifaReteIvaDbl)
                HcolPropiedades.Add(ObjTipoInterfazByt)
                HcolPropiedades.Add(ObjIdMedioPagoDefectoByt)
                HcolPropiedades.Add(ObjTipoTerceroCajaByt)
                HcolPropiedades.Add(ObjTotalAreaCopropDec)
                HcolPropiedades.Add(ObjTotalAreaPondDec)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    Friend Property EnuEstadoAplicacion As EnuEstadoAplicacionDef = EnuEstadoAplicacionDef.None
    Friend Property BlnImportarFacturas As Boolean = True
    Friend Shared ReadOnly Property ObjCentroUtilidad As ClsCentroUtilidad
        Get
            Return GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual
        End Get
    End Property
    Friend ReadOnly Property ObjAnoActual As ClsAno
        Get
            If MobjAnoActual Is Nothing Then
                Dim lblnAnoNoCerrado = False
                For Each lobjAno As ClsAno In ColAnos
                    If Not lobjAno.ObjEstaCerradoAnoBln.ObjValorPro Then
                        MobjAnoActual = lobjAno
                        lblnAnoNoCerrado = True
                        Exit For
                    End If
                Next
                If Not lblnAnoNoCerrado AndAlso McolAnos.Count > 0 Then
                    MobjAnoActual = McolAnos(McolAnos.Count)
                End If
            End If
            Return MobjAnoActual
        End Get
    End Property
    Friend ReadOnly Property ObjAno(ashrIdAno As Short) As ClsAno
        Get
            Dim lobjAno As ClsAno = Nothing
            If ColAnos.Contains(ashrIdAno.ToString) Then
                lobjAno = ColAnos(ashrIdAno.ToString)
            End If
            Return lobjAno
        End Get
    End Property
    ''' <summary>
    ''' Instancia y devuelve un objeto Servicio de tipo único de acuerdo a los valore de los argumentos.
    ''' </summary>
    ''' <param name="ashrIdAno">Año del servicio</param>
    ''' <param name="ashrIdServicio">Id. del Servicio</param>
    Friend ReadOnly Property ObjServicio(ashrIdAno As Short, ashrIdServicio As Short) As ClsServicio
        Get
            Dim lobjAno As ClsAno = Nothing
            Dim lobjServicio As ClsServicio
            If ashrIdAno > 0 Then
                lobjAno = ColAnos(ashrIdAno.ToString)
            End If
            lobjServicio = New ClsServicio(lobjAno, EnuModoInstanciaObjDef.enuUnico)
            Dim lobjValorLlave() As Object = {GshrIdCarpeta, GshrIdCentroUtil, ashrIdAno, ashrIdServicio}
            lobjServicio.SAbra(lobjValorLlave)
            Return lobjServicio
        End Get
    End Property
    Friend ReadOnly Property ObjProveedorEFac As ClsProveedorEFac
        Get
            If MobjProveedorEFac Is Nothing Then
                MobjProveedorEFac = ObjIdProvEFacByt.ObjProvEFac
            End If
            Return MobjProveedorEFac
        End Get
    End Property
    Friend ReadOnly Property EnuEstadoInstalacion As EnuEstadoInstalacion
        Get
            Return MenuEstadoInstalacion
        End Get
    End Property
    Friend ReadOnly Property StrPrefijoResStr As String
        Get
            Return FstrPrefijoDoc(EnuIdDocumentoDef.EnuFacturaVenta)
        End Get
    End Property
    Friend ReadOnly Property BlnEFacAutorizado As Boolean
        Get
            Dim lblnEFacAuto As Boolean = ObjIdProvEFacByt.ObjValorPro > EnuProveedorEFac.None AndAlso
                    ObjAutorizaEFacBln.ObjValorPro
            Return lblnEFacAuto
        End Get
    End Property
#End Region
#End Region

#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SVacie()
        MyBase.SVacie()
        McolAgrupadoresServicios = Nothing
        McolAnos = Nothing
        McolModulos = Nothing
        McolSectores = Nothing
        McolServiciosPer = Nothing
        McolDocumentos = Nothing
        McolFechasServicio = Nothing
        MobjDocumento = Nothing
        MobjAnoActual = Nothing
        MdtbAgrupadoresServicios = Nothing
        MdtbServiciosPer = Nothing
        MobjProveedorEFac = Nothing
    End Sub
    Protected Overrides Sub SModifique()
        Dim lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            MyBase.SModifique()
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                If ObjIdProvEFacByt.ObjValorPro > EnuProveedorEFac.None Then
                    Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil,
                            ObjIdProvEFacByt.ObjValorPro}
                    ObjProveedorEFac.SAbra(lobjValorLlave)
                    If Not ObjProveedorEFac.BlnExiste Then
                        ObjProveedorEFac.SNormaliceEstado(False)
                        ObjProveedorEFac.SCreeObj(lobjValorLlave)
                        ObjProveedorEFac.ObjIdProveedorEFacEnt.ObjValorPro = ObjIdProvEFacByt.ObjValorPro
                    Else
                        ObjProveedorEFac.SModifique()
                    End If
                End If
            End If
            lblnNoHayError = True
        Catch ex As PanDatException
            Throw
        Catch ex As PanLException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                GobjPanDat.SControleProcesoObj(False)
            Else
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        Dim lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            GobjPanDat.SInicialiceTransaccion()
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                ObjIdCarpetaCentroUtilOriCopShr.ObjValorPro = GshrIdCarpeta
                ObjIdCentroUtil_OriCopShr.ObjValorPro = GshrIdCentroUtil
                ObjTotalAreaCopropDec.ObjValorPro = 0
                SCreeDocumentos()
                ClsPanorama.SActualiceCol(McolDocumentos)
            End If
            If BlnEFacAutorizado Then
                ObjProveedorEFac.SActualice(ablnExigeRequeridos)
            End If
            Dim lstrPieFacRes = FstrPieFacturaRes()
            If Not String.IsNullOrEmpty(lstrPieFacRes) Then
                ObjPieFacturaUnoStr.ObjValorPro = lstrPieFacRes
            End If
            MyBase.SActualice(ablnExigeRequeridos)
            lblnNoHayError = True
        Catch ex As PanDatException
            Throw
        Catch ex As PanLException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                GobjAdministrador.BlnNotificacionSonora = ObjNotificacionesSonorasBln.ObjValorPro
                GobjPanDat.SConfirmeTransaccion()
                GobjPanDat.SControleProcesoObj(False)
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    Protected Overrides Sub SInicialiceObj()
        ObjFechaResolucionFactDtm.ObjValorPro = GCDTMFECHANULA
        ObjFechaVenceResolFactDtm.ObjValorPro = GCDTMFECHANULA
        ObjFechaUltCausacionGralDtm.ObjValorPro = GCDTMFECHANULA
        ObjFechaResolucionContDtm.ObjValorPro = GCDTMFECHANULA
        ObjIdCtaDescuentosPPStr.ObjValorPro = String.Empty
        ObjIdCtaImptosAsumidosStr.ObjValorPro = String.Empty
        ObjIdCtaReteFuenteStr.ObjValorPro = String.Empty
        ObjIdCtaReteIcaStr.ObjValorPro = String.Empty
        ObjIdCtaReteIvaStr.ObjValorPro = String.Empty
        ObjIdCtaIngPorIdentificarStr.ObjValorPro = String.Empty
        ObjAutorizaEFacBln.ObjValorPro = False
        ObjIdProvEFacByt.ObjValorPro = 0
        ObjRangoFraFinEnt.ObjValorPro = 0
        ObjRangoFraIniEnt.ObjValorPro = 0
        ObjRangoFraConIniEnt.ObjValorPro = 0
        ObjRangoFraConFinEnt.ObjValorPro = 0
        ObjPrefijoFactContStr.ObjValorPro = String.Empty
        ObjNumeroResolContiStr.ObjValorPro = String.Empty
        ObjPrefijoFactContStr.ObjValorPro = String.Empty
        ObjTipoInterfazByt.ObjValorPro = EnuTipoInterfazDef.None
        ObjTipoTerceroCajaByt.ObjValorPro = EnuTipoTerceroCajaDef.None
        ObjTarifaReteIvaDbl.ObjValorPro = 0
        ObjIdAppContableByt.ObjValorPro = 0
        ObjPieFacturaDosStr.ObjValorPro = String.Empty
        ObjPieFacturaUnoStr.ObjValorPro = String.Empty
        ObjFirmaRCeMail.ObjValorPro = False
        McolAnos = Nothing
        McolModulos = Nothing
        McolSectores = Nothing
        McolServiciosPer = Nothing
        ObjIdProveedorEFacEnt.ObjValorPro = 0
        ObjIdUsuarioProvEFacStr.ObjValorPro = String.Empty
        ObjIdCarpetaEFacShr.ObjValorPro = GshrIdCarpeta
        ObjIdCentroUtilEFacShr.ObjValorPro = GshrIdCentroUtil
        ObjURLStr.ObjValorPro = String.Empty
        ObjSubirFacBln.ObjValorPro = False
    End Sub
    Friend Overrides Sub SRefresqueObj()
        Dim ldtbTablaColeccion = ClsPanorama.FdtbDataTable(ClsCentroUtilOriCop.SstrNombreTabla,
                {"*"}, Nothing, ClsOrionCop.StrFiltroUbicacion)
        DtbTablaColeccion = ldtbTablaColeccion
        If DtbTablaColeccion.Rows.Count > 0 Then
            DrwRegistroActual = DtbTablaColeccion.Rows(0)
            SVacie()
            SLeaValores(True)
            GobjAdministrador.BlnNotificacionSonora = ObjNotificacionesSonorasBln.ObjValorPro
        Else
            EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando
            MyBase.SRefresqueObj()
        End If
    End Sub
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ""
        End Get
    End Property
    Friend Overrides Function FblnNotificaOk(aenuIdMensNot As EnuIdMens) As Boolean
        Dim lblnNotOk = False
        If aenuIdMensNot = EnuIdMens.EnuMensInicio Then
            SVerifiqueApp(False, True)
            If MenuEstadoInstalacion <> EnuEstadoInstalacion.Todos OrElse
                EnuEstadoAplicacion <> EnuEstadoAplicacionDef.EnuNormal Then
                lblnNotOk = False
            Else
                lblnNotOk = True
            End If
        End If
        If aenuIdMensNot = EnuIdMens.EnuServicioMalParam Then
            lblnNotOk = FblnSerBienParam()
        ElseIf aenuIdMensNot = EnuIdMens.None Then
            lblnNotOk = True
        End If
        Return lblnNotOk
    End Function
#End Region

#Region "Manejo Años"
    Friend ReadOnly Property ColAnos As Collection
        Get
            Static lblnCargandoCol As Boolean = False
            If McolAnos Is Nothing OrElse McolAnos.Count = 0 Then
                If Not lblnCargandoCol Then
                    lblnCargandoCol = True
                    McolAnos = New Collection
                    Dim ldtbAnos = FdtbAnos()
                    If ldtbAnos.Rows.Count > 0 Then
                        For Each ldrwAno As DataRow In ldtbAnos.Rows
                            Dim lobjAno As New ClsAno(Me, ldrwAno)
                            lobjAno.SLeaValores(True)
                            McolAnos.Add(lobjAno, lobjAno.ObjIdAnoShr.ToString)
                        Next
                    End If
                    lblnCargandoCol = False
                End If
            End If
            Return McolAnos
        End Get
    End Property
    Friend Function SCreeAno(ablnInicioMesActual As Boolean, adblIncrementoCA As Double,
            aenuTipoCalCuota As EnuTipoBaseCalculo) As ClsAno
        Dim lblnNoHayError = False
        Dim lobjAno As ClsAno
        Try
            GobjPanDat.SControleProcesoObj(True)
            GobjPanDat.SInicialiceTransaccion()
            Dim ldtbAnos = FdtbAnos()
            Dim ldrwNuevoAno = ldtbAnos.NewRow
            If ldtbAnos.Rows.Count = 0 Then
                lobjAno = FobjPrimerAno(ldrwNuevoAno, ablnInicioMesActual)
                If (Date.Today.Month = 1 AndAlso ablnInicioMesActual) OrElse
                        (Date.Today.Month = 2 AndAlso Not ablnInicioMesActual) Then
                    Dim ldrwSiguienteAnoAct As DataRow = ldtbAnos.NewRow()
                    Dim lobjSiguienteAno = FobjAnoSiguiente(ldrwSiguienteAnoAct, 0,
                            EnuTipoBaseCalculo.EnuCoeficientePro)
                End If
            Else
                lobjAno = FobjAnoSiguiente(ldrwNuevoAno, adblIncrementoCA,
                        aenuTipoCalCuota)
            End If
            lblnNoHayError = True
        Catch ex As PanDatException
            Throw
        Catch ex As ErrorInesperadoPanLException
            Throw
        Catch ex As PanLException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                McolAnos = Nothing
                GobjPanDat.SConfirmeTransaccion()
                GobjPanDat.SControleProcesoObj(False)
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
        Return lobjAno
    End Function
    ''' <summary>
    ''' Crea el primer Año de Orión y lo adiciona a la colección de años. 
    ''' </summary>
    ''' <param name="ablnEsAnoAnteriorAlPrimero">Indica si se debe crear el año anterior al actual, que debe ser
    ''' el único año creado.</param>
    ''' <remarks></remarks>
    ''' <summary>
    ''' Indica si el año es el ultimo año en las colección de años de la Copropiedad
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function FobjPrimerAno(adrwNuevoAno As DataRow, ablnInicioMesActual As Boolean) As ClsAno
        Dim lobjNuevoAno As New ClsAno(Me, adrwNuevoAno)
        Dim lshrIdAno = 0S, lstrPeriodoIni = String.Empty
        If ablnInicioMesActual Then
            If Date.Today.Month = 1 Then
                lshrIdAno = Year(Date.Today) - 1
            Else
                lshrIdAno = Year(Date.Today)
            End If
        Else
            If Date.Today.Month <= 2 Then
                lshrIdAno = Year(Date.Today) - 1
            Else
                lshrIdAno = Year(Date.Today)
            End If
        End If
        With lobjNuevoAno
            .SCreeObj(Nothing)
            .ObjIdAnoShr.ObjValorPro = lshrIdAno
            .ObjTipoCalculoCuotaByt.ObjValorPro = EnuTipoBaseCalculo.EnuCoeficientePro
            .ObjModuloPorServicioBln.ObjValorPro = True
            .BlnInicioMesActual = ablnInicioMesActual
            .SActualice(True)
        End With
        Return lobjNuevoAno
    End Function
    Private Function FobjAnoSiguiente(adrwNuevoAno As DataRow, adblIncremento As Double,
            aenuTipoCalCuota As EnuTipoBaseCalculo) As ClsAno
        Dim lobjNuevoAno As New ClsAno(Me, adrwNuevoAno)
        Dim lobjAnoAnteriorAEste As ClsAno = ColAnos(ColAnos.Count)
        Dim lshrIdNuevoAno As Short = lobjAnoAnteriorAEste.ObjIdAnoShr.ObjValorPro + 1
        Dim ldecValorPresRef = lobjAnoAnteriorAEste.ObjValorPres_AnoDec.ObjValorPro
        Dim ldecValorPres As Decimal = Math.Round(ldecValorPresRef * (1 + adblIncremento), 0)
        With lobjNuevoAno
            .SCreeObj(Nothing)
            .ObjIdAnoShr.ObjValorPro = lshrIdNuevoAno
            .ObjAplicaDsctoPPBln.ObjValorPro = lobjAnoAnteriorAEste.ObjAplicaDsctoPPBln.ObjValorPro
            .ObjDiasParaDsctoPPShr.ObjValorPro =
                    lobjAnoAnteriorAEste.ObjDiasParaDsctoPPShr.ObjValorPro
            .ObjModuloPorServicioBln.ObjValorPro =
                    lobjAnoAnteriorAEste.ObjModuloPorServicioBln.ObjValorPro
            .ObjTipoDsctoPPByt.ObjValorPro = lobjAnoAnteriorAEste.ObjTipoDsctoPPByt.ObjValorPro
            If GblnOK Then
                .ObjValorPres_AnoDec.ObjValorPro = ldecValorPres
            Else
                .ObjValorPres_AnoDec.ObjValorPro = 0D
            End If
            .ObjTipoCalculoCuotaByt.ObjValorPro = aenuTipoCalCuota
            If lobjAnoAnteriorAEste.ColServiciosAno.Count > 0 Then
                SAdicioneServiciosAno(lobjAnoAnteriorAEste, lobjNuevoAno, adblIncremento)
            End If
            .SActualice(True)
        End With
        If ldecValorPres > 0 Then
            Dim lstrMens = String.Empty
            If aenuTipoCalCuota = EnuTipoBaseCalculo.EnuCuotaAnterior Then
                ClsCalculosServicios.SCreeCuotasBaseAnoAnt(lobjAnoAnteriorAEste.ObjIdAnoShr.
                        ObjValorPro, adblIncremento)
            Else
                Dim lblnCreo = lobjNuevoAno.FblnCalculoCuotasAdmin(lstrMens)
            End If
        End If
        Return lobjNuevoAno
    End Function
    Friend Function FblnEsElUltimoAno(ashrIdAno As Short)
        Dim lblnEs = False
        If ColAnos.Count > 0 Then
            Dim lobjUltAno As ClsAno = ColAnos(ColAnos.Count)
            If Not IsNothing(lobjUltAno) Then
                lblnEs = (ashrIdAno = lobjUltAno.ObjIdAnoShr.ObjValorPro)
            End If
        End If
        Return lblnEs
    End Function
    Friend Function FblnEsElPrimerAno(ashrIdAno As Short)
        Dim lblnEs = False
        If ColAnos.Count > 0 Then
            Dim lobjPrimerAno As ClsAno = ColAnos(1)
            If Not IsNothing(lobjPrimerAno) Then
                lblnEs = (ashrIdAno = lobjPrimerAno.ObjIdAnoShr.ObjValorPro)
            End If
        End If
        Return lblnEs
    End Function
    ''' <summary>
    ''' Indica si el año actual es el primer año en el cual se van a generar facturas
    ''' </summary>
    ''' <returns></returns>
    Friend Function FblnAnoEsElPrimero(ashrIdAno As Short) As Boolean
        Dim lblnEsElPrimero = False
        Dim lobjAnoAnterior As ClsAno
        If ObjAnoActual IsNot Nothing Then
            If ColAnos.Contains((ashrIdAno - 1).ToString()) Then
                lobjAnoAnterior = ColAnos((ashrIdAno - 1).ToString())
                lblnEsElPrimero = lobjAnoAnterior.ObjValorPres_AnoDec.ObjValorPro = 0
            Else
                lblnEsElPrimero = True
            End If
        End If
        Return lblnEsElPrimero
    End Function
    Private Function FdtbAnos() As DataTable
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion
        Dim ldtbAnos = ClsPanorama.FdtbDataTable(ClsAno.SstrNombreTabla, {"*"},
                    {{StrCampoCarpeta, "ASC"},
                    {StrCampoCentroUtil, "ASC"},
                    {ClsIdAnoShr.SstrNombreCampoBd, "ASC"}}, lstrFiltro)
        Return ldtbAnos
    End Function
    Friend Function FblnExisteServicioAno(astrKeyServicio As String) As Boolean
        Dim lstrPartes() As String = astrKeyServicio.Split(",")
        Dim lstrIdano As String = lstrPartes(0)
        Dim lobjAno As ClsAno
        If ColAnos.Contains(lstrIdano) Then
            lobjAno = ColAnos(lstrIdano)
        Else
            Throw New ErrorInesperadoPanLException("El año no existe!")
        End If
        Return lobjAno.ColServiciosAno.Contains(astrKeyServicio)
    End Function
    Friend Function FblnPuedeCrearSerAnual(ashrIdAno As Short) As Boolean
        Dim lblnPuede As Boolean
        Dim lobjAno As ClsAno = ColAnos(ashrIdAno.ToString)
        lblnPuede = lobjAno.ColServiciosAno.Count > 0
        If Not lblnPuede Then
            Dim lshrIdAnoAnt = ashrIdAno - 1
            If ColAnos.Contains(lshrIdAnoAnt) Then
                lobjAno = ColAnos(lshrIdAnoAnt.ToString)
                lblnPuede = lobjAno.ColServiciosAno.Count > 0 AndAlso
                        lobjAno.ObjEstaCerradoAnoBln.ObjValorPro
            Else
                lblnPuede = True
            End If
        End If
        Return lblnPuede
    End Function
    Private Shared Sub SAdicioneServiciosAno(aobjAnoAnterior As ClsAno,
            aobjAnoNuevo As ClsAno, adblIncrementoCA As Double)
        Dim lcolServiciosAnoAnte As Collection = aobjAnoAnterior.ColServiciosAno
        For Each lobjServicioAnoAnt As ClsServicio In lcolServiciosAnoAnte
            If Not lobjServicioAnoAnt.ObjEsAjusteBln.ObjValorPro Then
                Dim lobjSerNuevoAno As ClsServicio =
                        aobjAnoNuevo.FobjNuevoServicioAno(False)
                Dim lstrNuevoPerIni As String =
                    CType(aobjAnoNuevo.ObjIdAnoShr.ObjValorPro, String) & "01"
                With lobjSerNuevoAno
                    .ObjCodigoCuentaCrStr.ObjValorPro =
                            lobjServicioAnoAnt.ObjCodigoCuentaCrStr.ObjValorPro
                    .ObjModoCausaInteresesByt.ObjValorPro =
                            lobjServicioAnoAnt.ObjModoCausaInteresesByt.ObjValorPro
                    .ObjDiasVencimientoShr.ObjValorPro =
                            lobjServicioAnoAnt.ObjDiasVencimientoShr.ObjValorPro
                    .ObjDiasGraciaShr.ObjValorPro =
                            lobjServicioAnoAnt.ObjDiasGraciaShr.ObjValorPro
                    .ObjVenceFinMesBln.ObjValorPro =
                            lobjServicioAnoAnt.ObjVenceFinMesBln.ObjValorPro
                    .ObjGraciaFinMesBln.ObjValorPro =
                            lobjServicioAnoAnt.ObjGraciaFinMesBln.ObjValorPro
                    .ObjCodigoCuentaDbStr.ObjValorPro =
                            lobjServicioAnoAnt.ObjCodigoCuentaDbStr.ObjValorPro
                    .ObjCodigoCuentaIvaStr.ObjValorPro =
                            lobjServicioAnoAnt.ObjCodigoCuentaIvaStr.ObjValorPro
                    .ObjCodigoCuentaDevStr.ObjValorPro =
                            lobjServicioAnoAnt.ObjCodigoCuentaDevStr.ObjValorPro
                    .ObjCodigoCuentaMoraStr.ObjValorPro =
                            lobjServicioAnoAnt.ObjCodigoCuentaMoraStr.ObjValorPro
                    .ObjBaseMinimaReteFuenteDec.ObjValorPro =
                            lobjServicioAnoAnt.ObjBaseMinimaReteFuenteDec.ObjValorPro
                    .ObjBaseMinimaReteIcaDec.ObjValorPro =
                            lobjServicioAnoAnt.ObjBaseMinimaReteIcaDec.ObjValorPro
                    .ObjEsServicioIdBln.ObjValorPro = False
                    .ObjEsExcluidoIvaBln.ObjValorPro =
                            lobjServicioAnoAnt.ObjEsExcluidoIvaBln.ObjValorPro
                    .ObjEsFactProgramableBln.ObjValorPro =
                            lobjServicioAnoAnt.ObjEsFactProgramableBln.ObjValorPro
                    .ObjEstaGenaradaProgramBln.ObjValorPro = False
                    .ObjGeneraProgramBln.ObjValorPro = True
                    .ObjTipoBaseCalculoByt.ObjValorPro =
                            aobjAnoNuevo.ObjTipoCalculoCuotaByt.ObjValorPro
                    .ObjNombreServicioStr.ObjValorPro =
                            lobjServicioAnoAnt.ObjNombreServicioStr.ObjValorPro
                    .ObjTarifaIvaDbl.ObjValorPro =
                            lobjServicioAnoAnt.ObjTarifaIvaDbl.ObjValorPro
                    .ObjTarifaRetFteDbl.ObjValorPro =
                            lobjServicioAnoAnt.ObjTarifaRetFteDbl.ObjValorPro
                    .ObjTarifaRetIcaDbl.ObjValorPro =
                            lobjServicioAnoAnt.ObjTarifaRetIcaDbl.ObjValorPro
                    .ObjPeriodoInicioStr.ObjValorPro = lstrNuevoPerIni
                    .ObjCantPeriodos_ServicioShr.ObjValorPro = 12
                    .ObjIdServicioShr.ObjValorPro =
                            lobjServicioAnoAnt.ObjIdServicioShr.ObjValorPro
                    .ObjIdTipoTerCtaCrSerByt.ObjValorPro =
                            lobjServicioAnoAnt.ObjIdTipoTerCtaCrSerByt.ObjValorPro
                    .ObjIdTerceroCtaCrDbl.ObjValorPro =
                            lobjServicioAnoAnt.ObjIdTerceroCtaCrDbl.ObjValorPro
                    .ObjConceptoServicioStr.ObjValorPro =
                            lobjServicioAnoAnt.ObjConceptoServicioStr.ObjValorPro
                End With
                aobjAnoNuevo.ColServiciosAno.Add(lobjSerNuevoAno)
                If lobjServicioAnoAnt.BlnEsCuotaAdministracion Then
                    lobjSerNuevoAno.SAdicioneModulosServicio(lobjServicioAnoAnt, adblIncrementoCA)
                End If
            End If
        Next
    End Sub
#End Region

#Region "Manejo Servicios Permanentes"
    Friend ReadOnly Property ColServiciosPer As Collection
        Get
            If IsNothing(McolServiciosPer) Then
                McolServiciosPer = New Collection
                SCargueDtbServiciosPer()
                If MdtbServiciosPer.Rows.Count > 0 Then
                    Dim ldrwServiciosPer() As DataRow = MdtbServiciosPer.Select
                    For Each ldrwServicioPer As DataRow In ldrwServiciosPer
                        Dim lobjServicioPer As New ClsServicio(Me, ldrwServicioPer)
                        lobjServicioPer.SLeaValores(True)
                        Dim lstrKey = lobjServicioPer.ObjIdAno_ServicioShr.ToString & "," &
                                lobjServicioPer.ObjIdServicioShr.ToString
                        McolServiciosPer.Add(lobjServicioPer, lstrKey)
                    Next
                End If
            End If
            Return McolServiciosPer
        End Get
    End Property
    Friend Function FblnExisteNombreServicio(astrNombreServicio As String) As Boolean
        SCargueDtbServiciosPer()
        Dim lstrFiltro = ClsNombreServicioStr.SstrNombreCampoBd & " = '" & astrNombreServicio & "'"
        Dim ldrwServicios() As DataRow = MdtbServiciosPer.Select(lstrFiltro)
        Return (ldrwServicios.Length > 0)
    End Function
    ''' <summary>
    ''' Devuelve el nombre del Servicio Permanente identificado con "ashrIdServicio".
    ''' </summary>
    ''' <param name="ashrIdServicio">Identidad del servicio del cual se devolverá el nombre.</param>
    ''' <returns>Nombre del servicio (String)</returns>
    ''' <remarks>Si no existe el Servicio se devuelve una cadena vacia.</remarks>
    Friend Function FstrNombreServicio(ashrIdServicio As Short) As String
        Dim lstrNombreServicio = String.Empty
        For Each lobjServicio As ClsServicio In ColServiciosPer
            If lobjServicio.ObjIdServicioShr.ObjValorPro = ashrIdServicio Then
                lstrNombreServicio = lobjServicio.ObjNombreServicioStr.ObjValorPro
                Exit For
            End If
        Next
        Return lstrNombreServicio
    End Function
    ''' <summary>
    ''' Devuelve el nombre del Servicio Anual del año "ashrIdAno" e identificado con "ashrIdServicio".
    ''' </summary>
    ''' <param name="ashrIdAno">Identidad del año al cual pertenece el servicio.</param>
    ''' <param name="ashrIdServicio">Identidad del servicio del cual se devolverá el nombre.</param>
    ''' <returns>Nombre del servicio (String)</returns>
    ''' <remarks>Si no existe el Servicio se devuelve una cadena vacia.</remarks>
    Friend Function FstrNombreServicio(ashrIdAno As Short, ashrIdServicio As Short) As String
        Dim lstrNombreServicio As String
        If ashrIdAno = 0 Then
            If ashrIdServicio = "999" Then
                lstrNombreServicio = "Intereses de mora"
            Else
                lstrNombreServicio = FstrNombreServicio(ashrIdServicio)
            End If
        Else
            Dim lobjAno As ClsAno = ColAnos(ashrIdAno.ToString)
            Dim lstrKey = ashrIdAno.ToString & "," & ashrIdServicio.ToString
            Dim lobjServicio As ClsServicio = lobjAno.ColServiciosAno(lstrKey)
            lstrNombreServicio = lobjServicio.ObjNombreServicioStr.ObjValorPro
        End If
        Return lstrNombreServicio
    End Function
    ''' <summary>
    ''' Devuelve el Id. del Servicio que tiene como nombre el argumento "astrNombreServicio"
    ''' </summary>
    ''' <param name="astrNombreServicio">Nombre del servicio del cual se devuelve su Identificación.</param>
    ''' <returns></returns>
    ''' <remarks>Si no existe el Servicio se devuelve cero.</remarks>
    Friend Function FshrIdServicio(astrNombreServicio As String, ablnNombre As Boolean) As Short
        Dim lshrIdServicio As Short = 0
        For Each lobjServicio As ClsServicio In ColServiciosPer
            If ablnNombre Then
                If lobjServicio.ObjNombreServicioStr.ObjValorPro = astrNombreServicio Then
                    lshrIdServicio = lobjServicio.ObjIdServicioShr.ObjValorPro
                    Exit For
                End If
            Else
                If lobjServicio.ObjConceptoServicioStr.ObjValorPro = astrNombreServicio Then
                    lshrIdServicio = lobjServicio.ObjIdServicioShr.ObjValorPro
                    Exit For
                End If
            End If
        Next
        Return lshrIdServicio
    End Function
    Friend Function FblnHayOtroServicioId(ashrIdServicio As Short) As Boolean
        Dim lblnHaySerId = False
        For Each lobjServicio As ClsServicio In ColServiciosPer
            If lobjServicio.ObjEsServicioIdBln.ObjValorPro Then
                lblnHaySerId = lobjServicio.ObjIdServicioShr.ObjValorPro <> ashrIdServicio
                If lblnHaySerId Then Exit For
            End If
        Next
        Return lblnHaySerId
    End Function
    Friend Function FblnHayServicioId() As Boolean
        Dim lblnHaySerId = False
        For Each lobjServicio As ClsServicio In ColServiciosPer
            lblnHaySerId = lobjServicio.ObjEsServicioIdBln.ObjValorPro
            If lblnHaySerId Then Exit For
        Next
        Return lblnHaySerId
    End Function
    Friend Function FobjServicioId() As ClsServicio
        Dim lobjServicio As ClsServicio = Nothing
        For Each lobjSer As ClsServicio In ColServiciosPer
            If lobjSer.ObjEsServicioIdBln.ObjValorPro Then
                lobjServicio = lobjSer
                Exit For
            End If
        Next
        Return lobjServicio
    End Function
    Private Sub SCargueDtbServiciosPer()
        If IsNothing(MdtbServiciosPer) Then
            Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsIdAno_ServicioShr.SstrNombreCampoBd & " = 0"
            MdtbServiciosPer = ClsPanorama.FdtbDataTable(ClsServicio.SstrNombreTabla, {"*"},
                    {{StrCampoCarpeta, "ASC"}, {StrCampoCentroUtil, "ASC"},
                     {ClsIdServicioShr.SstrNombreCampoBd, "ASC"}, {ClsIdAno_ServicioShr.SstrNombreCampoBd, "ASC"}},
                    lstrFiltro)
        End If
    End Sub
#End Region

#Region "Manejo Modulos de Contribucion"
    Friend ReadOnly Property ColModulos As Collection
        Get
            If IsNothing(McolModulos) Then
                McolModulos = New Collection
                Dim ldtbMod = FdtbModulos()
                If ldtbMod.Rows.Count > 0 Then
                    For Each ldrwModulo As DataRow In ldtbMod.Rows
                        Dim lobjModulo As New ClsModuloContribucion(Me, ldrwModulo)
                        lobjModulo.SLeaValores(True)
                        McolModulos.Add(lobjModulo, lobjModulo.ObjIdModuloShr.ToString)
                    Next
                End If
            End If
            Return McolModulos
        End Get
    End Property
    Friend Function FdtbModulos() As DataTable
        Dim lstrCamposSelect() As String = {StrCampoCarpeta,
                StrCampoCentroUtil,
                ClsIdModuloShr.SstrNombreCampoBd,
                ClsNombreModuloStr.SstrNombreCampoBd,
                ClsContribuyeCuotaAdminBln.SstrNombreCampoBd}
        Dim lstrIndice(,) As String = {{StrCampoCarpeta, "ASC"},
                {StrCampoCentroUtil, "ASC"},
                {ClsIdModuloShr.SstrNombreCampoBd, "ASC"}}
        Dim ldtbModulos = ClsPanorama.FdtbDataTable(ClsModuloContribucion.SstrNombreTabla,
                lstrCamposSelect, lstrIndice, ClsOrionCop.StrFiltroUbicacion)
        Return ldtbModulos
    End Function
    ''' <summary>
    ''' Calcula y devuelve el total de la base de participación con el factor de ponderación
    ''' aplicado
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function FdblTotalBasePart(aenuTipoBase As EnuTipoBaseCalculo) As Double
        Dim ldblTotalBasePart = 0.0
        For Each lobjModuloContr As ClsModuloContribucion In ColModulos
            If lobjModuloContr.ObjContribuyeCuotaAdminBln.ObjValorPro Then
                ldblTotalBasePart += lobjModuloContr.FdblBaseTotalParticipaModulo(aenuTipoBase)
            End If
        Next
        Return ldblTotalBasePart
    End Function
    ''' <summary>
    ''' Indica si el módulo de contribución identificado por el argumento "ashrIdModulo" contribuye
    ''' con un servicio cuota de administración e indica ademas en el argumento "ablnSoloAUnServicio" 
    ''' si contribuye con un solo servicio de administración.
    ''' </summary>
    ''' <param name="ashrIdModulo">Identificador del módulo de contribución</param>
    ''' <param name="ablnSoloAUnaCuota">Indica si solo contribuye con un solo servicio</param>
    Private Function FblnModuloContribuyeCuotaAdmin(ashrIdModulo As Short,
            ByRef ablnSoloAUnServicio As Boolean)
        Dim lblnContr As Boolean, i = 0
        lblnContr = GobjParametros.FblnPerActEsDicPrimerAno
        If Not lblnContr Then
            For Each lobjServicio As ClsServicio In ObjAnoActual.ColServiciosAno
                If lobjServicio.BlnEsCuotaAdministracion AndAlso
                    Not lobjServicio.ObjEsAjusteBln.ObjValorPro Then
                    If lobjServicio.ColModulosServicio.Contains(ashrIdModulo.ToString) Then
                        i += 1
                    End If
                End If
            Next
        Else
            i += 1
        End If
        lblnContr = i > 0
        ablnSoloAUnServicio = i = 1
        Return lblnContr
    End Function
#End Region

#Region "Manejo Sectores"
    Friend ReadOnly Property ColSectores As Collection
        Get
            If IsNothing(McolSectores) Then
                McolSectores = New Collection
                Dim ldtbSectores As DataTable = FdtbSectores()
                Dim ldrwSectores() As DataRow = ldtbSectores.Select
                For Each ldrwSector As DataRow In ldrwSectores
                    Dim lobjSector As New ClsSector(Me, ldrwSector)
                    lobjSector.SLeaValores(True)
                    McolSectores.Add(lobjSector, lobjSector.ObjIdSectorShr.ToString)
                Next
            End If
            Return McolSectores
        End Get
    End Property
    Friend Function FshrIdSector(astrNombreSector As String) As Short
        Dim ldtbSectores = FdtbSectores()
        Dim lstrFiltro As String = ClsNombreSectorStr.SstrNombreCampoBd & " = '" &
                astrNombreSector & "'"
        Dim ldrwSectores As DataRow() = ldtbSectores.Select(lstrFiltro)
        Return ClsPanorama.FobjValorCampo(ldrwSectores(0)(ClsIdSectorShr.SstrNombreCampoBd),
                EnuTipoValor.enuShort)
    End Function
    Friend Function FdtbSectores() As DataTable
        Dim lstrCamposSelect() As String = {StrCampoCarpeta,
                StrCampoCentroUtil,
                ClsIdSectorShr.SstrNombreCampoBd,
                ClsNombreSectorStr.SstrNombreCampoBd,
                "'0' AS Area",
                ClsDctoProntoPago_SecDbl.SstrNombreCampoBd,
                "'' AS DsctoForma"}
        Dim lstrIndice(,) As String = {{StrCampoCarpeta, "ASC"},
                                         {StrCampoCentroUtil, "ASC"},
                                         {ClsIdSectorShr.SstrNombreCampoBd, "ASC"}}
        Dim ldtbSectores = ClsPanorama.FdtbDataTable(ClsSector.SstrNombreTabla, lstrCamposSelect,
                    lstrIndice, ClsOrionCop.StrFiltroUbicacion)
        SComplementeDtbSectores(ldtbSectores)
        Return ldtbSectores
    End Function
    Private Sub SComplementeDtbSectores(adtbSectores As DataTable)
        If (Not IsNothing(ObjAnoActual)) AndAlso adtbSectores.Rows.Count > 0 Then
            Dim lenuTipoDsctoPP = ObjAnoActual.ObjTipoDsctoPPByt.ObjValorPro
            Dim lstrValorDscto As String, ldblDscto As Double
            Dim lobjSector As New ClsSector(EnuModoInstanciaObjDef.enuUnico)
            Dim lshrIdSector As Short, lobjValorLlave As Object(), ldblAreaSector As Double
            For Each ldrwSector As DataRow In adtbSectores.Rows
                ldblDscto = ClsPanorama.FobjValorCampo(ldrwSector("DctoProntoPago"),
                        EnuTipoValor.enuDouble)
                If lenuTipoDsctoPP = EnuTipoDsctoPP.EnuProcentaje Then
                    lstrValorDscto = Format(ldblDscto, "p")
                ElseIf lenuTipoDsctoPP = EnuTipoDsctoPP.EnuValorFijo Then
                    lstrValorDscto = Format(ldblDscto, "c")
                Else
                    lstrValorDscto = Format(0, "c")
                End If
                ldrwSector("DsctoForma") = lstrValorDscto
                lshrIdSector = ClsPanorama.FobjValorCampo(ldrwSector(ClsIdSectorShr.SstrNombreCampoBd),
                        EnuTipoValor.enuShort)
                lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lshrIdSector}
                lobjSector.SAbra(lobjValorLlave)
                ldblAreaSector = lobjSector.FdblTotalAreaPrediosSector()
                ldrwSector("Area") = Format(Math.Round(ldblAreaSector, 3), "#0.00")
            Next
        End If
    End Sub
#End Region

#Region "Manejo Tasas Mora"
    Friend Function FdtbTasasMora() As DataTable
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion
        Dim lstrCamposSelect = {StrCampoCarpeta,
                                StrCampoCentroUtil,
                                ClsOrdinalTasaMoraEnt.SstrNombreCampoBd,
                                ClsFechaDesdeTasaMoraDtm.SstrNombreCampoBd,
                                ClsFechaHastaTasaMoraDtm.SstrNombreCampoBd,
                                ClsTasaMoraDbl.SstrNombreCampoBd,
                                "TasaMora/12 AS TasaMoraMes"}
        Dim lstrIndice(,) As String = {{StrCampoCarpeta, "ASC"},
                                    {StrCampoCentroUtil, "ASC"},
                                    {ClsOrdinalTasaMoraEnt.SstrNombreCampoBd, "ASC"}}
        Dim ldtbTasasMora = ClsPanorama.FdtbDataTable(ClsTasaMora.SstrNombreTabla, lstrCamposSelect,
                lstrIndice, lstrFiltro)
        If ldtbTasasMora.Rows.Count > 0 Then
            Dim lstrFecha = Format(Today, GCSTRFMTFECHASIMPLE)
            Dim ldrwTasMor As DataRow = ldtbTasasMora(ldtbTasasMora.Rows.Count - 1)
            ldrwTasMor("FechaHasta") = lstrFecha
        End If
        Return ldtbTasasMora
    End Function
    Friend Function FdblTasaMoraFecha(adtmFecha As Date) As Double
        Dim ldblTasa = 0.0
        Dim ldtbTasasMora = FdtbTasasMora()
        Dim ldtmFechaDesde As Date, ldtmFechaHasta As Date
        adtmFecha = adtmFecha.AddDays(-1)
        For Each ldrwTasaMora As DataRow In ldtbTasasMora.Rows
            ldtmFechaDesde = ClsPanorama.FobjValorCampo(ldrwTasaMora(
                    ClsFechaDesdeTasaMoraDtm.SstrNombreCampoBd), EnuTipoValor.enuDate)
            ldtmFechaHasta = ClsPanorama.FobjValorCampo(ldrwTasaMora(
                    ClsFechaHastaTasaMoraDtm.SstrNombreCampoBd), EnuTipoValor.enuDate)
            If adtmFecha >= ldtmFechaDesde AndAlso adtmFecha <= ldtmFechaHasta Then
                ldblTasa = ClsPanorama.FobjValorCampo(ldrwTasaMora(
                        ClsTasaMoraDbl.SstrNombreCampoBd), EnuTipoValor.enuDouble)
                Exit For
            End If
        Next
        Return ldblTasa
    End Function
    Friend Function FdblTasaMoraPeriodoActual() As Double
        Dim ldtmFechaFinPeriodo = ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
        Dim ldblTasa = FdblTasaMoraFecha(ldtmFechaFinPeriodo)
        Return ldblTasa
    End Function
#End Region

#Region "Manejo Cuentas Bancos"
    Friend Function FcolCuentasBanco() As Collection
        Dim lcolCuentasBanco As New Collection
        Dim ldtbCtasBco = FdtbCuentasBanco()
        If ldtbCtasBco.Rows.Count > 0 Then
            Dim ldrwCuentasBanco() As DataRow = ldtbCtasBco.Select
            For Each ldrwCuentaBanco As DataRow In ldrwCuentasBanco
                Dim lobjCuentaBanco As New ClsCuentaBanco(Me, ldrwCuentaBanco)
                lobjCuentaBanco.SLeaValores(True)
                lcolCuentasBanco.Add(lobjCuentaBanco,
                            lobjCuentaBanco.ObjIdCuentaBancoShr.ToString)
            Next
        End If
        Return lcolCuentasBanco
    End Function
    Friend Function FblnEsCuentaBanco(astrIdCta As String) As Boolean
        Dim lblEsCtaBco = False
        Dim lcolCtasBanco = FcolCuentasBanco()
        For Each lobjCtaBco As ClsCuentaBanco In lcolCtasBanco
            If lobjCtaBco.ObjIdCtaContabilidadStr.ObjValorPro = astrIdCta Then
                lblEsCtaBco = True
                Exit For
            End If
        Next
        Return lblEsCtaBco
    End Function
    Friend Function FdtbCuentasBanco() As DataTable
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion
        Dim lstrCamposSelect = {"*"}
        Dim lstrIndice(,) As String = {{StrCampoCarpeta, "ASC"},
                {StrCampoCentroUtil, "ASC"},
                {ClsIdCuentaBancoShr.SstrNombreCampoBd, "ASC"}}
        Dim ldtbCtasBco = ClsPanorama.FdtbDataTable(ClsCuentaBanco.SstrNombreTabla, lstrCamposSelect,
                        lstrIndice, lstrFiltro)
        Return ldtbCtasBco
    End Function
#End Region

#Region "Manejo Documentos"
    Friend ReadOnly Property ColDocumentos As Collection
        Get
            If IsNothing(McolDocumentos) Then
                McolDocumentos = New Collection
                Dim ldtbDocs = FdtbDocumentos()
                For Each ldrwDocumento As DataRow In ldtbDocs.Rows
                    Dim lobjDocumento As New ClsDocumento(Me, ldrwDocumento) With {
                        .EnuPermisosObj =
                            GobjPanorama.FenuTipoPermisos(
                            EnuIdClasesPanDef.EnuDocumentoContabilidad)
                    }
                    lobjDocumento.SLeaValores(True)
                    McolDocumentos.Add(lobjDocumento, lobjDocumento.ObjIdDocumentoEnt.ToString)
                Next
            End If
            Return McolDocumentos
        End Get
    End Property
    Friend ReadOnly Property ObjDocumento(aenuIdDocumento As EnuIdDocumentoDef) As ClsDocumento
        Get
            ObjDocumento = Nothing
            Dim lshrIdDocumento As Short = CType(aenuIdDocumento, Short)
            If ColDocumentos.Count > 0 Then
                If ColDocumentos.Contains(lshrIdDocumento.ToString) Then
                    ObjDocumento = McolDocumentos(lshrIdDocumento.ToString)
                End If
            Else
                Dim ldtbDocs = FdtbDocumentos()
                Dim ldrwDoc = ldtbDocs.NewRow
                ObjDocumento = New ClsDocumento(Me, ldrwDoc)
            End If
            Return ObjDocumento
        End Get
    End Property
    Friend Function FobjNuevoDocumento() As ClsDocumento
        Dim lobjNuevoDoc As ClsDocumento = Nothing
        Dim lenuIdSigDoc = FenuIdSiguienteDocumento()
        If lenuIdSigDoc > EnuIdDocumentoDef.None + 1 Then
            Dim ldtbDocs = FdtbDocumentos()
            Dim ldrwNewDoc As DataRow = ldtbDocs.NewRow
            lobjNuevoDoc = New ClsDocumento(Me, ldrwNewDoc)
            With lobjNuevoDoc
                .SCreeObj(Nothing)
                .ObjIdCarpeta_DocShr.ObjValorPro = GshrIdCarpeta
                .ObjIdCentroUtil_DocShr.ObjValorPro = GshrIdCentroUtil
                .ObjIdDocumentoEnt.ObjValorPro = CType(lenuIdSigDoc, Short)
                .ObjNombre_DocStr.ObjValorPro = FstrNombreDoc(CType(lenuIdSigDoc, Short))
                .ObjNumeroInicial_DocEnt.ObjValorPro = 0
                .ObjTipoDocumentoStr.ObjValorPro = String.Empty
            End With
        End If
        Return lobjNuevoDoc
    End Function
    Friend Function FobjNuevoDocumento(aenuIdDocumento As EnuIdDocumentoDef) As ClsDocumento
        Dim lobjNuevoDoc As ClsDocumento = Nothing
        Dim ldtbDocs = FdtbDocumentos()
        Dim ldrwNewDoc As DataRow = ldtbDocs.NewRow
        lobjNuevoDoc = New ClsDocumento(Me, ldrwNewDoc)
        With lobjNuevoDoc
            .SCreeObj(Nothing)
            .ObjIdCarpeta_DocShr.ObjValorPro = GshrIdCarpeta
            .ObjIdCentroUtil_DocShr.ObjValorPro = GshrIdCentroUtil
            .ObjIdDocumentoEnt.ObjValorPro = CType(aenuIdDocumento, Short)
            .ObjNombre_DocStr.ObjValorPro = FstrNombreDoc(CType(aenuIdDocumento, Short))
            .ObjPrefijo_DocStr.ObjValorPro = String.Empty
            .ObjNumeroInicial_DocEnt.ObjValorPro = 0
            .ObjTipoDocumentoStr.ObjValorPro = String.Empty
        End With
        Return lobjNuevoDoc
    End Function
    Friend Sub SAdicioneNuevoDocumento(aobjNuevoDocumento As ClsDocumento)
        If Not IsNothing(aobjNuevoDocumento) Then
            If Not ColDocumentos.Contains(aobjNuevoDocumento.ObjIdDocumentoEnt.ToString) Then
                McolDocumentos.Add(aobjNuevoDocumento, aobjNuevoDocumento.ObjIdDocumentoEnt.ToString)
            End If
        End If
    End Sub
    Friend Function FblnHayDocumentosPorCrear() As Boolean
        Return (ColDocumentos.Count < 9)
    End Function
    Friend Function FblnDocumentosPorDefinir() As Boolean
        Dim lblnDocPorDocPorDefinir As Boolean, lobjDoc As ClsDocumento
        If ObjTipoInterfazByt.ObjValorPro = EnuTipoInterfazDef.None Then
            lblnDocPorDocPorDefinir = False
        ElseIf ObjTipoInterfazByt.ObjValorPro = EnuTipoInterfazDef.EnuPorComprobante Then
            lobjDoc = ColDocumentos(9)
            lblnDocPorDocPorDefinir = String.IsNullOrEmpty(lobjDoc.ObjTipoDocumentoStr.ToString())
        Else
            For Each lobjDocu As ClsDocumento In ColDocumentos
                If lobjDocu.ObjIdDocumentoEnt.ObjValorPro <= 8 Then
                    lblnDocPorDocPorDefinir =
                            String.IsNullOrEmpty(lobjDocu.ObjTipoDocumentoStr.ToString())
                    If lblnDocPorDocPorDefinir Then Exit For
                End If
            Next
        End If
        Return lblnDocPorDocPorDefinir
    End Function
    Private Function FenuIdSiguienteDocumento() As EnuIdDocumentoDef
        Dim lenuIdSiguiDoc = EnuIdDocumentoDef.None
        McolDocumentos = Nothing
        McolDocumentos = ColDocumentos
        If McolDocumentos.Count <= EnuIdDocumentoDef.EnuNotaAjuste + 1 Then
            lenuIdSiguiDoc = McolDocumentos.Count + 1
            FobjNuevoDocumento()
        End If
        Return lenuIdSiguiDoc
    End Function
    Friend Shared Function FstrNombreDoc(aenuIdDocumento As EnuIdDocumentoDef) As String
        Dim lstrNombreDoc = String.Empty
        Select Case aenuIdDocumento
            Case 1
                If GobjParametros.BlnEFacAutorizado Then
                    lstrNombreDoc = "FACTURA DE VENTA"
                Else
                    lstrNombreDoc = "CUENTA DE COBRO"
                End If
            Case 2
                lstrNombreDoc = "RECIBO DE CAJA"
            Case 3
                lstrNombreDoc = "NOTA APLICACION ANTICIPO"
            Case 4
                lstrNombreDoc = "NOTA DEBITO INTERESES DE MORA"
            Case 5
                lstrNombreDoc = "NOTA CREDITO"
            Case 6
                lstrNombreDoc = "NOTA REINTEGRO ANTICIPO"
            Case 7
                lstrNombreDoc = "NOTA REVERSION CREDITO"
            Case 8
                lstrNombreDoc = "NOTA AJUSTE CUOTA ADMINISTRACION"
            Case 9
                lstrNombreDoc = "COMPROBANTE INTERFAZ CONTABLE"
        End Select
        Return lstrNombreDoc
    End Function
    Friend Function FdtbDocumentos() As DataTable
        Dim ldtbDocumentos = ClsPanorama.FdtbDataTable(ClsDocumento.SstrNombreTabla, {"*"},
                {{ClsIdDocumentoEnt.SstrNombreCampoBd, ""}}, ClsOrionCop.StrFiltroUbicacion)
        Return ldtbDocumentos
    End Function
    Friend Function FstrPrefijoDoc(aenuIdDocumento As EnuIdDocumentoDef) As String
        Dim lstrPref = String.Empty
        If ColDocumentos.Count > 0 Then
            If ColDocumentos.Contains((CType(aenuIdDocumento, Integer).ToString)) Then
                MobjDocumento = ColDocumentos(CType(aenuIdDocumento, Integer).ToString)
                lstrPref = MobjDocumento.ObjPrefijo_DocStr.ObjValorPro
            End If
        End If
        Return lstrPref
    End Function
    Friend Function FentNumeracionInicialDoc(aenuIdDocumento As EnuIdDocumentoDef) As Integer
        MobjDocumento = ColDocumentos(CType(aenuIdDocumento, Integer).ToString)
        Return MobjDocumento.ObjNumeroInicial_DocEnt.ObjValorPro
    End Function
    Friend Function FstrTipoDoc(aenuIdDocumento As EnuIdDocumentoDef) As String
        MobjDocumento = ColDocumentos(CType(aenuIdDocumento, Integer).ToString)
        Return MobjDocumento.ObjTipoDocumentoStr.ObjValorPro
    End Function
    Private Sub SCreeDocumentos()
        McolDocumentos = ColDocumentos
        If McolDocumentos.Count = 0 Then
            Dim lenuIdSiguiDoc = EnuIdDocumentoDef.None
            Do Until lenuIdSiguiDoc = EnuIdDocumentoDef.EnuComprobanteInterfaz
                lenuIdSiguiDoc += 1
                Dim lobjDoc = FobjNuevoDocumento(lenuIdSiguiDoc)
                If lenuIdSiguiDoc = EnuIdDocumentoDef.EnuComprobanteInterfaz Then
                    lobjDoc.ObjTipoDocumentoStr.ObjValorPro = "CI"
                End If
                SAdicioneNuevoDocumento(lobjDoc)
            Loop
        End If
    End Sub
#End Region

#Region "Métodos Generales"
    Friend Sub SInicialicePorDefecto()
        ObjDiasParaPersuasivoShr.ObjValorPro = 90
        ObjDiasParaPrejuridicoShr.ObjValorPro = 180
        ObjDiasParaJuridicoShr.ObjValorPro = 270
        ObjDiasParaPerdidaShr.ObjValorPro = 720
        ObjBaseRedondeoCPByt.ObjValorPro = 1
        ObjBaseRedondeoIntMoraDbl.ObjValorPro = 100
        ObjBaseRedondeoGeneralDbl.ObjValorPro = 1000
        ObjIdMedioPagoDefectoByt.ObjValorPro = EnuTipoMedioPagoDef.EnuTransferencia
        ObjPlazoDefectoFacManualShr.ObjValorPro = 30
        ObjInformaSaldoTotalDespuesRCBln.ObjValorPro = True
        ObjNotificacionesSonorasBln.ObjValorPro = True
        ObjConsolidaItemsFacBln.ObjValorPro = False
        SAsigneVlaloresDefectoPro()
    End Sub
    ''' <summary>
    ''' Devuleve el servicio identificado por el paramentro lstrKey el cual es un string compuesto por el
    ''' año del servicio y el id del servicio separados por una coma: (IdAno,IdServicio). Si el servicio
    ''' no existe devuelve Nothing
    ''' </summary>
    ''' <param name="astrKeyServicio">Llave que identifica al servicio en las colecciones de servicios</param>
    Friend Function FobjServicio(astrKeyServicio As String) As ClsServicio
        Dim lobjServicio As ClsServicio = Nothing
        Dim lstrPartesKey() = Split(astrKeyServicio, ",")
        Dim lshrIdAno = CType(lstrPartesKey(0), Short)
        If lshrIdAno = 0 Then
            If ColServiciosPer.Contains(astrKeyServicio) Then
                lobjServicio = ColServiciosPer(astrKeyServicio)
            End If
        Else
            Dim lobjAno As ClsAno = ObjAno(lshrIdAno)
            If Not IsNothing(lobjAno) Then
                If lobjAno.ColServiciosAno.Contains(astrKeyServicio) Then
                    lobjServicio = lobjAno.ColServiciosAno(astrKeyServicio)
                End If
            End If
        End If
        Return lobjServicio
    End Function
    Friend Sub SCierrePeriodo()
        Dim lstrPeriodoActual = ObjAnoActual.StrIdPeriodoActual
        Dim lstrPeriodoHoy = ClsOrionCop.FstrPeriodoDeFecha(Date.Today)
        Dim lstrSiguientePeriodo = ClsOrionCop.FstrPeriodoFinal(lstrPeriodoActual, 1)
        If lstrSiguientePeriodo <= lstrPeriodoHoy OrElse (Now >
                ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo AndAlso Now.Hour > 12) Then
            Dim lblnCerrarAno = (ObjAnoActual.ObjPeriodoActual.ObjIdPeriodoShr.ObjValorPro = 12)
            ObjAnoActual.SCierrePeriodoActual()
            If lblnCerrarAno Then
                With ObjAnoActual
                    If .EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                        .EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                    End If
                    .ObjEstaCerradoAnoBln.ObjValorPro = True
                    .SActualice(True)
                End With
            End If
        End If
        McolAnos = Nothing
        Me.SRefresqueObj()
    End Sub
    Friend Sub SAbraPeriodoAnterior()
        Dim lshrIdAnoActual As Short = ObjAnoActual.ObjIdAnoShr.ObjValorPro
        Dim lstrPeriodoActual = ObjAnoActual.StrIdPeriodoActual
        Dim lstrPeriodoAnterior = ClsOrionCop.FstrPeriodoFinal(lstrPeriodoActual, -1)
        Dim lshrIdAnoPerAnt As Short = lstrPeriodoAnterior.Substring(0, 4)
        Dim lstrIdPerAnt As String = lstrPeriodoAnterior.Substring(4, 2)
        Dim lblnAbrirAnoAnterior = lshrIdAnoPerAnt < lshrIdAnoActual
        Dim lobjPerAnt As ClsPeriodo = Nothing
        If lblnAbrirAnoAnterior Then
            If GobjParametros.ColAnos.Contains(lshrIdAnoPerAnt.ToString) Then
                Dim lobjAnoAnt As ClsAno = GobjParametros.ColAnos(lshrIdAnoPerAnt.ToString)
                With lobjAnoAnt
                    If .EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                        .EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                    End If
                    .ObjEstaCerradoAnoBln.ObjValorPro = False
                    .SActualice(True)
                End With
                lobjPerAnt = lobjAnoAnt.ColPeriodos(lstrIdPerAnt)
                McolAnos = Nothing
                SRefresqueObj()
            End If
        Else
            lobjPerAnt = ObjAnoActual.ColPeriodos(lstrIdPerAnt)
        End If
        lobjPerAnt.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
        lobjPerAnt.ObjEstaCerradoPeriodoBln.ObjValorPro = False
        lobjPerAnt.SActualice(True)
    End Sub
    Friend Sub SRegistreFechaUltCausa(adtmfechaUltCausa As Date)
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
            EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
            ObjFechaUltCausacionGralDtm.ObjValorPro = adtmfechaUltCausa
            SActualice(True)
        End If
    End Sub
    Friend Sub SRegistreFechaFacturación()
        Dim lobjPeriodoActual = ObjAnoActual.ObjPeriodoActual
        With lobjPeriodoActual
            If .EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                .EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
            End If
            .ObjFechaFacturacionPeriodoDtm.ObjValorPro = Date.Today
            .SActualice(True)
        End With
    End Sub
    ''' <summary>
    ''' Indica si los sectores, modulos de contribución y los servicios estan bien parametrizados.
    ''' </summary>
    ''' <returns>Boolean</returns>
    ''' <remarks>Todos los sectorees deben contribuir con la cuota de administración así su tasa de
    ''' participación o ponderación sea cero. Asi mismo, todos los modulos de contribución marcados como 
    ''' contribuyentes de la cuota de administración, deben contribuir con ella</remarks>
    Friend Function FblnParaCuotaAdminOk(ByRef astrMens As String) As Boolean
        ' Verificar que los modulo marcados como contribuyentes de la cuota de administración 
        ' realmente lo esten haciendo y solo lo hagan a un servicio
        Dim lblnParaOk = False, lblnModulosOk = True, lblnSoloAUnServicio = True
        Dim i = 0
        If ObjAnoActual IsNot Nothing Then
            For Each lobjModuloCont As ClsModuloContribucion In ColModulos
                If lobjModuloCont.ObjContribuyeCuotaAdminBln.ObjValorPro Then
                    i += 1
                    If Not FblnModuloContribuyeCuotaAdmin(
                                lobjModuloCont.ObjIdModuloShr.ObjValorPro,
                                lblnSoloAUnServicio) Then
                        astrMens = "El Módulo de Contribución '" &
                            lobjModuloCont.ObjNombreModuloStr.ObjValorPro &
                            "' no está contribuyendo con la Cuota de Administración!"
                        lblnModulosOk = False
                    Else
                        If Not lblnSoloAUnServicio Then
                            astrMens = "El Módulo de Contribución '" &
                                lobjModuloCont.ObjNombreModuloStr.ObjValorPro &
                                "' esta contribuyendo con más de " &
                                "un Servicio Cuota de Administración!"
                            lblnModulosOk = False
                        End If
                    End If
                End If
                If Not lblnModulosOk Then Exit For
            Next
            If i = 0 Then
                lblnModulosOk = False
                astrMens = "No hay Módulos de Contribución marcados para contribuir " &
                        "con la Cuota de Administración!"
            End If
        End If
        If lblnModulosOk Then
            lblnParaOk = FblnTodosSectoresCont(astrMens)
        End If
        If Not IsNothing(ObjAnoActual) Then
            If Not FblnPerActEsDicPrimerAno() AndAlso lblnParaOk AndAlso
                    ObjAnoActual.ObjModuloPorServicioBln.ObjValorPro Then
                lblnParaOk = FblnHayUnModuloPorServicio(astrMens)
                If lblnParaOk Then

                End If
            End If
        End If
        Return lblnParaOk
    End Function
    ''' <summary>
    ''' Indica si todos los sectores contribuyen con la cuota de administración
    ''' </summary>
    ''' <returns>Falso o Verdadero</returns>
    ''' <remarks>Todos los sectores tienen que contribuir con la cuota administración, así 
    ''' su ponderación en la participación sea el cero por ciento</remarks>
    Private Function FblnTodosSectoresCont(ByRef astrMens As String) As Boolean
        Dim lblnFactorMayorCero = False, lblnContr = True
        For Each lobjSector As ClsSector In ColSectores
            lblnContr = FblnSectorContribuye(lobjSector.ObjIdSectorShr.ObjValorPro,
                lblnFactorMayorCero)
            If Not lblnContr Then
                astrMens = "El Sector '" & lobjSector.ObjNombreSectorStr.ObjValorNuevo +
                        "' no contribuye con la Cuota de Administración!"
                Exit For
            End If
        Next
        Return lblnContr
    End Function
    ''' <summary>
    ''' Indica si el sector identificado con el argumento "ashrIdSector" contribuye al servicio
    ''' Cuota de administración y si la tasa de contribución es mayor a cero en el argumento
    ''' ablnTasaContribucionMayorcero
    ''' </summary>
    ''' <param name="ashrIdSector">Id del Sector</param>
    ''' <param name="ablnTasaContribucionMayorcero">Indica si el factor de contribución 
    ''' es mayor a cero</param>
    Friend Function FblnSectorContribuye(ashrIdSector As Short,
            ByRef ablnTasaContribucionMayorCero As Boolean) As Boolean
        Dim lblnContribuye = False, lblnContribuyeEsteModulo As Boolean
        Dim lobjSectorModulo As ClsSectorModulo
        For Each lobjModulo As ClsModuloContribucion In ColModulos
            If lobjModulo.ObjContribuyeCuotaAdminBln.ObjValorPro Then
                lblnContribuyeEsteModulo = lobjModulo.ColSectoresModulo.Contains(ashrIdSector.ToString)
                If Not lblnContribuye Then
                    lblnContribuye = lblnContribuyeEsteModulo
                End If
                If lblnContribuyeEsteModulo Then
                    lobjSectorModulo = lobjModulo.ColSectoresModulo(ashrIdSector.ToString)
                    ablnTasaContribucionMayorCero = (lobjSectorModulo.ObjTasaContribucionDbl.ObjValorPro > 0)
                    If ablnTasaContribucionMayorCero Then Exit For
                End If
            End If
        Next
        Return lblnContribuye
    End Function
    ''' <summary>
    ''' Indica si las cuotas de administración se calculan con base en el area ponderada de 
    ''' cada predio
    ''' </summary>
    ''' <returns></returns>
    Friend Function FblnAreaPonderada() As Boolean
        Dim lblnSi
        Dim lobjAno As ClsAno = ObjAnoActual
        lblnSi = lobjAno.FblnCalcularCuotaAdminPorCP
        Return lblnSi
    End Function
    Private Function FblnDsctoPPSectoresOk() As Boolean
        If ObjAnoActual IsNot Nothing Then
            If ObjAnoActual.ObjAplicaDsctoPPBln.ObjValorPro Then
                For Each lobjModuCont As ClsModuloContribucion In ColModulos
                    If lobjModuCont.ObjContribuyeCuotaAdminBln.ObjValorPro Then
                        If Not FblnDsctoPPSecModuloOK(lobjModuCont) Then
                            Return False
                        End If
                    End If
                Next
            End If
        End If
        Return True
    End Function
    ''' <summary>
    ''' Indica si un sector de un módulo que contribuye a la administración tiene definido
    ''' el descuento por pronto pago.
    ''' </summary>
    ''' <param name="aobjModu">Módulo que contribuye a la administración</param>
    ''' <returns></returns>
    Private Function FblnDsctoPPSecModuloOK(aobjModu As ClsModuloContribucion)
        Dim lobjSector As ClsSector, ldblDscto = 0.0
        For Each lobjSecMod As ClsSectorModulo In aobjModu.ColSectoresModulo
            If lobjSecMod.ObjTasaContribucionDbl.ObjValorPro > 0 Then
                lobjSector = lobjSecMod.ObjSector_SectorModulo
                ldblDscto += lobjSector.ObjDctoProntoPago_SecDbl.ObjValorPro
            End If
        Next
        Return ldblDscto <> 0
    End Function
    Private Shared Function FblnHayUnModuloPorServicio(ByRef astrMens As String) As Boolean
        Dim lblnHay = True
        If Not IsNothing(GobjParametros.ObjAnoActual) Then
            For Each lobjServicio As ClsServicio In GobjParametros.ObjAnoActual.ColServiciosAno
                If Not lobjServicio.ObjEsAjusteBln.ObjValorPro Then
                    lblnHay = (lobjServicio.ColModulosServicio.Count = 1)
                    If Not lblnHay Then
                        astrMens = "Hay al menos un Servicio Cuota de Administración " &
                                "con más de un Módulo de Contribución asociado!"
                        Exit For
                    End If
                End If
            Next
        Else
            lblnHay = False
        End If
        Return lblnHay
    End Function
    Friend Function FblnPerActEsDicPrimerAno()
        Dim lblnCierto = False
        If Not IsNothing(ObjAnoActual) Then
            Dim lshrIdAnoRef As Short = ObjAnoActual.ObjIdAnoShr.ObjValorPro
            If ColAnos.Count > 0 Then
                Dim lobjAno As ClsAno = ColAnos(1)
                If lobjAno.ObjIdAnoShr.ObjValorPro = lshrIdAnoRef Then
                    lblnCierto = GobjParametros.ObjFechaUltCausacionGralDtm.ObjValorPro =
                            GCDTMFECHANULA
                    lblnCierto = lblnCierto AndAlso
                            (lobjAno.ObjPeriodoActual.ObjIdPeriodoShr.ObjValorPro = 12)
                End If
            End If
        End If
        Return lblnCierto
    End Function
    ''' <summary>
    ''' Devuelve el pie factura construido a partir de la resolución de autorización de facturas 
    ''' </summary>
    ''' <returns></returns>
    Friend Function FstrPieFacturaRes() As String
        Dim lstrPieFac = String.Empty
        If Not String.IsNullOrEmpty(ObjNumeroResolFacturaStr.ToString) Then
            Dim lstrNroRes = ObjNumeroResolFacturaStr.ObjValorPro
            Dim lstrFechaRes = String.Empty, lstrFechaVenRes = String.Empty
            If ObjFechaResolucionFactDtm.ObjValorPro <> GCDTMFECHANULA Then
                lstrFechaRes = ClsPanorama.FstrFechaddmmaaaaSepSlash(
                        ObjFechaResolucionFactDtm.ObjValorPro)
            End If
            If ObjFechaVenceResolFactDtm.ObjValorPro <> GCDTMFECHANULA Then
                lstrFechaVenRes = ClsPanorama.FstrFechaddmmaaaaSepSlash(
                            ObjFechaVenceResolFactDtm.ObjValorPro)
            End If
            Dim lstrPref = FstrPrefijoDoc(EnuIdDocumentoDef.EnuFacturaVenta)
            Dim lstrIniRango = ObjRangoFraIniEnt.ToString
            Dim lstrFinRango = ObjRangoFraFinEnt.ToString
            If Not String.IsNullOrEmpty(lstrFechaRes) AndAlso
                        Not String.IsNullOrEmpty(lstrFechaVenRes) AndAlso
                        Not String.IsNullOrEmpty(lstrIniRango) AndAlso
                        Not String.IsNullOrEmpty(lstrFinRango) Then
                If String.IsNullOrEmpty(lstrPref) Then
                    lstrPieFac = "Habilita facturación Res. DIAN No. " & lstrNroRes &
                        " de fecha " & lstrFechaRes & ". Rango: del " &
                        lstrIniRango & " al " & lstrFinRango & "."
                Else
                    lstrPieFac = "Habilita facturación Res. DIAN No. " & lstrNroRes &
                        " de fecha " & lstrFechaRes & ". Rango: del " & lstrPref & "-" &
                        lstrIniRango & " al " & lstrPref & "-" & lstrFinRango & "."
                End If
            End If
        End If
        Return lstrPieFac
    End Function
    ''' <summary>
    ''' Devuelve el pie factura construido a partir de la resolución de autorización de facturas 
    ''' </summary>
    ''' <returns></returns>
    Friend Function FstrPieFacturaResCon() As String
        Dim lstrPieFac = String.Empty
        If Not String.IsNullOrEmpty(ObjNumeroResolContiStr.ToString) Then
            Dim lstrNroRes = ObjNumeroResolContiStr.ObjValorPro
            Dim lstrFechaRes = String.Empty
            If ObjFechaResolucionContDtm.ObjValorPro <> GCDTMFECHANULA Then
                lstrFechaRes = ClsPanorama.FstrFechaddmmaaaaSepSlash(
                        ObjFechaResolucionContDtm.ObjValorPro)
            End If
            Dim lstrPref = ObjPrefijoFactContStr.ObjValorPro
            Dim lstrIniRango = ObjRangoFraConIniEnt.ToString
            Dim lstrFinRango = ObjRangoFraConFinEnt.ToString
            If Not String.IsNullOrEmpty(lstrFechaRes) AndAlso
                        Not String.IsNullOrEmpty(lstrIniRango) AndAlso
                        Not String.IsNullOrEmpty(lstrFinRango) Then
                If String.IsNullOrEmpty(lstrPref) Then
                    lstrPieFac = "Habilita facturación Res. DIAN No. " & lstrNroRes &
                        " de fecha " & lstrFechaRes & ". Rango: del " &
                        lstrIniRango & " al " & lstrFinRango & "."
                Else
                    lstrPieFac = "Habilita facturación Res. DIAN No. " & lstrNroRes &
                        " de fecha " & lstrFechaRes & ". Rango: del " & lstrPref & "-" &
                        lstrIniRango & " al " & lstrPref & "-" & lstrFinRango & "."
                End If
            End If
        End If
        Return lstrPieFac
    End Function
    Friend Sub SActualicePieFactura()
        EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
        ObjPieFacturaUnoStr.ObjValorPro = FstrPieFacturaRes()
        SActualice(True)
    End Sub
    ''' <summary>
    ''' Indica si ya existe un concepto permanente con el concepto pasado en el argumento
    ''' </summary>
    ''' <returns></returns>
    Friend Function FblnExisteConcepto(astrKeySer As String, astrConcepto As String,
            ablnCreando As Boolean) As Boolean
        Dim lblnExiste = False, lstrKeySer As String
        If Not ablnCreando Then
            For Each lobjSer As ClsServicio In ColServiciosPer
                lstrKeySer = lobjSer.ObjIdAno_ServicioShr.ToString & "," &
                    lobjSer.ObjIdServicioShr.ToString
                If lstrKeySer <> astrKeySer Then
                    lblnExiste = lobjSer.ObjConceptoServicioStr.ObjValorPro = astrConcepto
                    If lblnExiste Then Exit For
                End If
            Next
        Else
            For Each lobjSer As ClsServicio In ColServiciosPer
                lblnExiste = lobjSer.ObjConceptoServicioStr.ObjValorPro = astrConcepto
                If lblnExiste Then Exit For
            Next
        End If
        Return lblnExiste
    End Function
    ' Fechas
    Friend ReadOnly Property ColFechasServicio() As Collection
        Get
            If McolFechasServicio Is Nothing Then
                McolFechasServicio = New Collection
                Dim lshrIdAno As Short, lshrIdSer As Short, lstrKeySer As String
                Dim ldtmFecFac As Date
                Dim lobjFechasSer As ClsFechasServicio, lobjServicio As ClsServicio
                Dim ldtbServPorFact = ClsOrionCop.FdtbServiciosPorFacturar(
                        EnuDestiItemProgramaFact.EnuTodos)
                For Each ldrwSerPorFac As DataRow In ldtbServPorFact.Rows
                    lshrIdAno = ClsPanorama.FobjValorCampo(ldrwSerPorFac(
                            ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd),
                            EnuTipoValor.enuShort)
                    lshrIdSer = ClsPanorama.FobjValorCampo(ldrwSerPorFac(
                            ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd),
                            EnuTipoValor.enuShort)
                    lstrKeySer = lshrIdAno.ToString() & "," & lshrIdSer.ToString()
                    lobjServicio = FobjServicio(lstrKeySer)
                    ldtmFecFac = lobjServicio.DtmFechaFacturacionPeriodoActual
                    If ldtmFecFac <= Date.Today Then
                        If McolFechasServicio.Count = 0 Then
                            lobjFechasSer = New ClsFechasServicio(ldtmFecFac)
                            lobjFechasSer.SAdicioneKeySer(lstrKeySer)
                            McolFechasServicio.Add(lobjFechasSer)
                        Else
                            If Not FblnExisteEnFechasSer(lobjServicio) Then
                                lobjFechasSer = New ClsFechasServicio(ldtmFecFac)
                                lobjFechasSer.SAdicioneKeySer(lstrKeySer)
                                McolFechasServicio.Add(lobjFechasSer)
                            End If
                        End If
                    End If
                Next
            End If
            Return McolFechasServicio
        End Get
    End Property
    Friend Function FarlFechasAFActurarHoy() As ArrayList
        Dim lshrIdAno As Short, lshrIdSer As Short, lstrKeySer As String
        Dim ldtmFecFac As Date, lobjServicio As ClsServicio
        Dim larlFecAFac As New ArrayList
        Dim ldtbServPorFact = ClsOrionCop.FdtbServiciosPorFacturar(
                        EnuDestiItemProgramaFact.EnuTodos)
        For Each ldrwSerPorFac As DataRow In ldtbServPorFact.Rows
            lshrIdAno = ClsPanorama.FobjValorCampo(ldrwSerPorFac(
                            ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd),
                            EnuTipoValor.enuShort)
            lshrIdSer = ClsPanorama.FobjValorCampo(ldrwSerPorFac(
                            ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd),
                            EnuTipoValor.enuShort)
            lstrKeySer = lshrIdAno.ToString() & "," & lshrIdSer.ToString()
            lobjServicio = FobjServicio(lstrKeySer)
            ldtmFecFac = lobjServicio.DtmFechaFacturacionPeriodoActual
            If ldtmFecFac <= Date.Today Then
                If Not larlFecAFac.Contains(ldtmFecFac) Then
                    larlFecAFac.Add(ldtmFecFac)
                End If
            End If
        Next
        larlFecAFac.Sort()
        Return larlFecAFac
    End Function
    Friend Sub SEstablezcaFechasServicios(adtmFecha As Date, ablnVence As Boolean)
        McolFechasServicio = Nothing
        If ablnVence Then
            For Each lobjSer As ClsServicio In ObjAnoActual.ColServiciosAno
                lobjSer.DtmFechaVencePeriActual = adtmFecha
            Next
            For Each lobjSer As ClsServicio In ColServiciosPer
                lobjSer.DtmFechaVencePeriActual = adtmFecha
            Next
        Else
            For Each lobjSer As ClsServicio In ObjAnoActual.ColServiciosAno
                lobjSer.DtmFechaGraciaPeriActual = adtmFecha
            Next
            For Each lobjSer As ClsServicio In ColServiciosPer
                lobjSer.DtmFechaGraciaPeriActual = adtmFecha
            Next
        End If
    End Sub
    Private Function FblnExisteEnFechasSer(aobjServicio As ClsServicio) As Boolean
        Dim ldtmFecFac = aobjServicio.DtmFechaFacturacionPeriodoActual
        Dim lstrKeySer As String
        Dim lblnExiste As Boolean
        For Each lobjFechasSer As ClsFechasServicio In McolFechasServicio
            lblnExiste = Not lobjFechasSer.FblnEsDistinta(ldtmFecFac)
            If lblnExiste Then
                lstrKeySer = aobjServicio.ObjIdAno_ServicioShr.ToString() & "," &
                    aobjServicio.ObjIdServicioShr.ToString
                lobjFechasSer.SAdicioneKeySer(lstrKeySer)
                Exit For
            End If
        Next
        Return lblnExiste
    End Function
    Friend Function FblnFechaFacturacionUnica(ByRef adtmFechaFacturacion As Date) As Boolean
        Dim ldtmFecFac As Date = GCDTMFECHANULA, lblnFechaUnica = True
        For Each lobjFechasSer As ClsFechasServicio In ColFechasServicio
            If ldtmFecFac = GCDTMFECHANULA Then
                ldtmFecFac = lobjFechasSer.DtmFechaFac
            Else
                lblnFechaUnica = ldtmFecFac = lobjFechasSer.DtmFechaFac
                If Not lblnFechaUnica Then Exit For
            End If
        Next
        If lblnFechaUnica Then
            adtmFechaFacturacion = ldtmFecFac
        Else
            adtmFechaFacturacion = GCDTMFECHANULA
        End If
        Return lblnFechaUnica
    End Function
    Friend Function FblnFechaVenceUnica(ByRef adtmFechaVence As Date) As Boolean
        Dim ldtmFecVen As Date = GCDTMFECHANULA, ldtmFecVenSer As Date
        Dim lblnFechaUnica As Boolean
        For Each lobjFechasSer As ClsFechasServicio In ColFechasServicio
            ldtmFecVenSer = FdtmFechaVence(lobjFechasSer)
            If ldtmFecVenSer = GCDTMFECHANULA Then
                ldtmFecVen = GCDTMFECHANULA
                Exit For
            Else
                If ldtmFecVen = GCDTMFECHANULA Then
                    ldtmFecVen = ldtmFecVenSer
                Else
                    If ldtmFecVen <> ldtmFecVenSer Then
                        ldtmFecVen = GCDTMFECHANULA
                        Exit For
                    End If
                End If
            End If
        Next
        adtmFechaVence = ldtmFecVen
        lblnFechaUnica = ldtmFecVen <> GCDTMFECHANULA
        Return lblnFechaUnica
    End Function
    ''' <summary>
    ''' Devuelve la fecha de vencimiento de los servicios que incluye la fecha de facturación incluidos
    ''' en al paramentro aobjFechasSer. Si no es la misma fecha para todos los servicios devuelve
    ''' la fecha GCDTMFECHANULA
    ''' </summary>
    ''' <param name="aobjFechasSer"></param>
    ''' <returns></returns>
    Private Function FdtmFechaVence(aobjFechasSer As ClsFechasServicio) As Date
        Dim lobjServicio As ClsServicio, ldtmFechaVence = GCDTMFECHANULA
        For Each lstrKeySer As String In aobjFechasSer.StrKeysSer
            lobjServicio = FobjServicio(lstrKeySer)
            If ldtmFechaVence = GCDTMFECHANULA Then
                ldtmFechaVence = lobjServicio.DtmFechaVencePeriActual
            Else
                If ldtmFechaVence <> lobjServicio.DtmFechaVencePeriActual Then
                    ldtmFechaVence = GCDTMFECHANULA
                    Exit For
                End If
            End If
        Next
        Return ldtmFechaVence
    End Function
    Friend Function FblnFechaGraciaUnica(ByRef adtmFechaGracia As Date) As Boolean
        Dim ldtmFecGra As Date = GCDTMFECHANULA, ldtmFecGraSer As Date
        Dim lblnFechaUnica As Boolean
        For Each lobjFechasSer As ClsFechasServicio In ColFechasServicio
            ldtmFecGraSer = FdtmFechaGracia(lobjFechasSer)
            If ldtmFecGraSer = GCDTMFECHANULA Then
                ldtmFecGra = GCDTMFECHANULA
                Exit For
            Else
                If ldtmFecGra = GCDTMFECHANULA Then
                    ldtmFecGra = ldtmFecGraSer
                Else
                    If ldtmFecGra <> ldtmFecGraSer Then
                        ldtmFecGra = GCDTMFECHANULA
                        Exit For
                    End If
                End If
            End If
        Next
        adtmFechaGracia = ldtmFecGra
        lblnFechaUnica = ldtmFecGra <> GCDTMFECHANULA
        Return lblnFechaUnica
    End Function
    ''' <summary>
    ''' Devuelve la fecha de gracia de los servicios que incluye la fecha de facturación incluidos
    ''' en al paramentro aobjFechasSer. Si no es la misma fecha para todos los servicios devuelve
    ''' la fecha GCDTMFECHANULA
    ''' </summary>
    ''' <param name="aobjFechasSer"></param>
    ''' <returns></returns>
    Private Function FdtmFechaGracia(aobjFechasSer As ClsFechasServicio) As Date
        Dim lobjServicio As ClsServicio, ldtmFechaGracia = GCDTMFECHANULA
        For Each lstrKeySer As String In aobjFechasSer.StrKeysSer
            lobjServicio = FobjServicio(lstrKeySer)
            If ldtmFechaGracia = GCDTMFECHANULA Then
                ldtmFechaGracia = lobjServicio.DtmFechaGraciaPeriActual
            Else
                If ldtmFechaGracia <> lobjServicio.DtmFechaGraciaPeriActual Then
                    ldtmFechaGracia = GCDTMFECHANULA
                    Exit For
                End If
            End If
        Next
        Return ldtmFechaGracia
    End Function
    Friend Function FblnFechaGraciaValida(aobjFechasFact As ClsFechasServicio,
            adtmFechaGracia As Date) As Boolean
        Dim lobjServicio As ClsServicio, lblnEsValida As Boolean
        For Each lstrKeySer As String In aobjFechasFact.StrKeysSer
            lobjServicio = FobjServicio(lstrKeySer)
            lblnEsValida = adtmFechaGracia >= lobjServicio.DtmFechaVencePeriActual
            If Not lblnEsValida Then Exit For
        Next
        Return lblnEsValida
    End Function
#End Region

#Region "Verifica Instalacion y Estado"
    Friend Function SVerifiqueApp(ablnNotifica As Boolean, ablnVerifiqueEstado As Boolean) As String
        Dim lstrMens = String.Empty
        If Not GobjParametros.ObjParametrizacionOkBln.ObjValorPro Then
            SVerifiqueInstalacion(lstrMens)
        Else
            MenuEstadoInstalacion = EnuEstadoInstalacion.Todos
        End If
        If MenuEstadoInstalacion = EnuEstadoInstalacion.Todos Then
            If BlnEFacAutorizado Then
                If Not GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro Then
                    lstrMens = "En Facturación Electrónica es necesario que la Opción: " & Chr(34) &
                            "Exige Fecha de Hoy en Fecha de Documentos" & Chr(34) & ", esté activa!"
                End If
            End If
            If String.IsNullOrEmpty(lstrMens) Then
                If ablnVerifiqueEstado Then
                    SVerifiqueEstado(lstrMens)
                End If
                Select Case EnuEstadoAplicacion
                    Case EnuEstadoAplicacionDef.EnuDocPorDefinir
                        lstrMens = "Es necesario definir los Documentos Contables"
                    Case EnuEstadoAplicacionDef.EnuCrearServicioAno
                        lstrMens = "Se deben crear los Servicios del Año actual!"
                    Case EnuEstadoAplicacionDef.EnuServicioNotOk
                        If String.IsNullOrEmpty(lstrMens) Then
                            lstrMens = "Hay al menos un Servicio Permanente que debe ser calculado!"
                        End If
                    Case EnuEstadoAplicacionDef.EnuSectoresSinDsctoPP
                        lstrMens = "Es necesario definir el Descuento por Pronto Pago para los Sectores!"
                    Case EnuEstadoAplicacionDef.EnuListoImportar
                        lstrMens = "El Sistema está listo para hacer la Importación inicial!"
                    Case EnuEstadoAplicacionDef.EnuListoImpNDb
                        lstrMens = "El Sistema está listo para la importación de las Notas por " &
                            "Intereses de Mora y/o Cierre del Mes!"
                    Case EnuEstadoAplicacionDef.EnuSinPresupuesto
                        lstrMens = "Se debe ingresar el Presupueso de Ingresos por Cuotas de Administarción!"
                    Case EnuEstadoAplicacionDef.EnuSinModulos
                        lstrMens = "Se deben parametrizar los Módulos que Contribuyen a los Servicios del Año!"
                    Case EnuEstadoAplicacionDef.EnuDebeAjustarCuotasAdmin
                        lstrMens = "Se debe generar el Retroactivo de las Cuotas de Administración!"
                    Case EnuEstadoAplicacionDef.enuDebeImportarAjuste
                        lstrMens = "Se deben importar las Cuotas del Retroactivo!"
                    Case EnuEstadoAplicacionDef.EnuHayItemsProgFactPorProcesar
                        If BlnEFacAutorizado Then
                            SVerifiqueResDian(lstrMens)
                            If Not String.IsNullOrEmpty(lstrMens) Then
                                lstrMens = lstrMens.Substring(0, 1).ToLower &
                                        lstrMens.Substring(1)
                                lstrMens = "Hay Pre-Facturas por generar" & " y " & lstrMens
                            End If
                        End If
                        If String.IsNullOrEmpty(lstrMens) Then
                            If BlnEFacAutorizado Then
                                lstrMens = "La Aplicación está lista para generar las Pre-Facturas!"
                            Else
                                lstrMens = "La Aplicación está lista para generar las Pre-Cuentas de Cobro!"
                            End If
                        End If
                    Case EnuEstadoAplicacionDef.EnuHayPrefacturas
                        SVerifiqueResDian(lstrMens)
                        If Not String.IsNullOrEmpty(lstrMens) Then
                            lstrMens = lstrMens.Substring(0, 1).ToLower & lstrMens.Substring(1)
                            lstrMens = "Se deben generar las Facturas definitivas" & " y " &
                                lstrMens
                        Else
                            If BlnEFacAutorizado Then
                                lstrMens = "Se deben generar las Facturas definitivas!"
                            Else
                                lstrMens = "Se deben generar las Cuentas de Cobro definitivas!"
                            End If
                        End If
                    Case EnuEstadoAplicacionDef.EnuCrearAno
                        Dim lstrAno = (ObjAnoActual.ObjIdAnoShr.ObjValorPro + 1).ToString()
                        lstrMens = "Se debe crear el Año " & lstrAno & "!"
                    Case EnuEstadoAplicacionDef.EnuCausarInt
                        SVerifiqueResDian(lstrMens)
                        If String.IsNullOrEmpty(lstrMens) Then
                            lstrMens = "Se deben causar los Intereses de Mora! "
                        End If
                    Case EnuEstadoAplicacionDef.EnuParaCierreMes
                        lstrMens = "El Sistema está listo para el Cierre de Mes!"
                    Case EnuEstadoAplicacionDef.EnuSinCalcAdmin
                        lstrMens = "Se deben calcular las Cuotas de Administración!"
                    Case EnuEstadoAplicacionDef.EnuFactFueraDeFecha
                        lstrMens = "Hay Facturas con fecha diferente a hoy sin procesar " &
                                "electrónicamente!"
                End Select
                If String.IsNullOrEmpty(lstrMens) Then
                    FblnServicioIdOk(lstrMens)
                End If
            End If
        End If
        If ablnNotifica Then
            If Not String.IsNullOrEmpty(lstrMens) Then
                SLevanteEventoNot(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            Else
                SLevanteEventoNot(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuOk)
            End If
        End If
        Return lstrMens
    End Function

    Friend Sub SVerifiqueInstalacion(ByRef astrMens As String)
        MenuEstadoInstalacion = EnuEstadoInstalacion.None
        GobjPanDat.SControleProcesoObj(True)
        Try
            SVerifiqueInstalacion_1(astrMens)
            SVerifiqueInstalacion_2(astrMens)
            ' Verifica que esten calculados los coeficientes de propiedad
            If MenuEstadoInstalacion And EnuEstadoInstalacion.Predios Then
                If FblnCPCalculados() Then
                    MenuEstadoInstalacion += EnuEstadoInstalacion.CPCalculados
                Else
                    astrMens = My.Resources.VerCoefProp
                End If
            End If
            ' Verifica que existan los propietarios
            If MenuEstadoInstalacion And EnuEstadoInstalacion.CPCalculados Then
                If FblnHayPropietarios() Then
                    MenuEstadoInstalacion += EnuEstadoInstalacion.Propietarios
                Else
                    astrMens = My.Resources.VerProp
                End If
            End If
            ' Verifica se hayan definido los Documentos
            If MenuEstadoInstalacion And EnuEstadoInstalacion.Propietarios Then
                If Not FblnHayDocumentosPorCrear() Then
                    If FblnDocumentosPorDefinir() Then
                        astrMens = My.Resources.VerDocsNoDef
                    Else
                        MenuEstadoInstalacion += EnuEstadoInstalacion.Docum
                    End If
                Else
                    astrMens = My.Resources.VerDocs
                End If
            End If
            ' Verifica que se haya creado el servicio de identificacion
            If MenuEstadoInstalacion And EnuEstadoInstalacion.Docum Then
                If FblnServicioIdOk(astrMens) Then
                    MenuEstadoInstalacion += EnuEstadoInstalacion.SerId
                End If
            End If
            If MenuEstadoInstalacion = EnuEstadoInstalacion.Todos Then
                SModifique()
                ObjParametrizacionOkBln.ObjValorPro = True
                SActualice(False)
                MenuEstadoInstalacion = EnuEstadoInstalacion.Todos
            End If
        Catch ex As ProveedorBdPanException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As PanLException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        End Try
        GobjPanDat.SControleProcesoObj(False)
    End Sub

    Private Sub SVerifiqueInstalacion_1(ByRef astrMens As String)
        ' Verifica la existencia de cuentas contables terminales.
        If FblnCuentasContOk() Then
            MenuEstadoInstalacion += EnuEstadoInstalacion.CuentasCont
        Else
            astrMens = My.Resources.VerCtasCont
        End If
        ' Verifica que las Opciones de la Copropiedad esten definidas
        If MenuEstadoInstalacion And EnuEstadoInstalacion.CuentasCont Then
            If Me.FblnEstanTodosOk Then
                MenuEstadoInstalacion += EnuEstadoInstalacion.OpcionesCentroUtil
            Else
                astrMens = My.Resources.VerCenUtil
            End If
        End If
        ' Verifica las Cuentas Bancarias
        If MenuEstadoInstalacion And EnuEstadoInstalacion.OpcionesCentroUtil Then
            If FblnCuentasBancosOk() Then
                MenuEstadoInstalacion += EnuEstadoInstalacion.CuentasBancos
            Else
                astrMens = My.Resources.VerBancos
            End If
        End If
        ' Verifica que haya Sectores
        If MenuEstadoInstalacion And EnuEstadoInstalacion.CuentasBancos Then
            If ColSectores.Count > 0 Then
                MenuEstadoInstalacion += EnuEstadoInstalacion.Sectores
            Else
                astrMens = My.Resources.VerSect
            End If
        End If
        ' Verifica que haya Modulos
        If MenuEstadoInstalacion And EnuEstadoInstalacion.Sectores Then
            If ColModulos.Count > 0 Then
                MenuEstadoInstalacion += EnuEstadoInstalacion.Modulos
            Else
                astrMens = My.Resources.VerModulos
            End If
        End If
        ' Verifica que uno de los modulos tenga al menos un sector contribuyente asignado
        If MenuEstadoInstalacion And EnuEstadoInstalacion.Modulos Then
            If FblnSectoresModulos() Then
                MenuEstadoInstalacion += EnuEstadoInstalacion.SectoresModulo
            Else
                astrMens = My.Resources.VerSectModulo
            End If
        End If
        'Verifica que haya Años
        If MenuEstadoInstalacion And EnuEstadoInstalacion.SectoresModulo Then
            If ColAnos.Count > 0 Then
                MenuEstadoInstalacion += EnuEstadoInstalacion.Anos
            Else
                astrMens = My.Resources.VerAnos
            End If
        End If
    End Sub

    Private Sub SVerifiqueInstalacion_2(ByRef astrMens As String)
        'Verifica que el año actual tenga al menos un Servicio creado
        If MenuEstadoInstalacion And EnuEstadoInstalacion.Anos Then
            MobjAnoActual = Nothing
            Dim lblnServicios = False
            If Not IsNothing(ObjAnoActual) Then
                lblnServicios = ObjAnoActual.ColServiciosAno.Count > 0
                If Not lblnServicios Then
                    astrMens = My.Resources.VerServicioAno & " " &
                            ObjAnoActual.ObjIdAnoShr.ToString() & "!"
                Else
                    lblnServicios = FblnParaCuotaAdminOk(astrMens)
                End If
                If lblnServicios Then
                    lblnServicios = ColServiciosPer.Count > 0
                    If Not lblnServicios Then
                        astrMens = My.Resources.VerServiciosPer
                    End If
                End If
                If lblnServicios Then
                    MenuEstadoInstalacion += EnuEstadoInstalacion.Servicios
                End If
            Else
                astrMens = My.Resources.VerAnoActual
            End If
        End If
        ' Verifica que al menos tenga una tasa de mora definida
        If MenuEstadoInstalacion And EnuEstadoInstalacion.Servicios Then
            If FblnHayTasaMora() Then
                MenuEstadoInstalacion += EnuEstadoInstalacion.TasasMora
            Else
                astrMens = My.Resources.VerTasasMora
            End If
        End If
        ' Verifica que haya Terceros
        If MenuEstadoInstalacion And EnuEstadoInstalacion.TasasMora Then
            If FblnTercerosCreados() Then
                MenuEstadoInstalacion += EnuEstadoInstalacion.Terceros
            Else
                astrMens = My.Resources.VerTerceros
            End If
        End If
        ' Verifica que haya Clientes
        If MenuEstadoInstalacion And EnuEstadoInstalacion.Terceros Then
            If FblnClientesCreados() Then
                MenuEstadoInstalacion += EnuEstadoInstalacion.Clientes
            Else
                astrMens = My.Resources.VerClientes
            End If
        End If
        ' Verifica que haya Predios
        If MenuEstadoInstalacion And EnuEstadoInstalacion.Clientes Then
            If FblnPrediosCreados() Then
                MenuEstadoInstalacion += EnuEstadoInstalacion.Predios
            Else
                astrMens = My.Resources.VerPredios
            End If
        End If
    End Sub

    Private Sub SVerifiqueEstado(ByRef astrMens As String)
        GobjPanDat.SControleProcesoObj(True)
        If BlnImportarFacturas Then
            BlnImportarFacturas = Not FblnHayFacturas()
        End If
        EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuNormal
        Select Case True
            Case FblnDocumentosPorDefinir()
                EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuDocPorDefinir
            Case Not ClsOrionCop.FblnEmailOk(astrMens)
                EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuEmailNoOk
            Case Not ObjAnoActual.ColServiciosAno.Count > 0
                EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuCrearServicioAno
            Case Not FblnParaCuotaAdminOk(astrMens)
                EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuParaSerAdminNotOK
            Case Not FblnServiciosOk(astrMens)
                EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuServicioNotOk
            Case FblnServiciosPerPorCalcular(astrMens)
                EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuServicioPorCal
            Case Not FblnDsctoPPSectoresOk()
                EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuSectoresSinDsctoPP
            Case BlnImportarFacturas
                EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuListoImportar
            Case FblbImportarNotasDbIni()
                EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuListoImpNDb
            Case ObjAnoActual.ObjValorPres_AnoDec.ObjValorPro = 0
                If FblnPerActEsDicPrimerAno() Then
                    If GobjParametros.ColAnos.Count = 1 Then
                        EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuCrearAno
                    Else
                        EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuParaCierreMes
                    End If
                ElseIf Not ClsOrionCop.FblnHacerCierreMes Then
                    If ObjAnoActual.ObjModuloPorServicioBln.ObjValorPro Then
                        EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuSinPresupuesto
                    Else
                        EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuSinModulos
                    End If
                Else
                    EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuParaCierreMes
                End If
            Case ClsOrionCop.FblnHayFacsFueraFecha()
                If Not GblnPosteando Then
                    EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuFactFueraDeFecha
                End If
            Case ClsOrionCop.FblnDocPorProcesarEFac()
                If Not GblnPosteando Then
                    EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuDocPorProEFac
                End If
            Case ClsOrionCop.FblnHayItemsPorFacturar AndAlso ClsOrionCop.FblnDebeCausarInt
                EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuCausarInt
            Case ClsOrionCop.FblnHayPrefacturas
                EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuHayPrefacturas
            Case ClsOrionCop.FblnHacerCierreMes
                If ObjAnoActual.ObjPeriodoActual.ObjIdPeriodoShr.ObjValorPro = 12 Then
                    Dim lobjUltimoAno As ClsAno = ColAnos(ColAnos.Count)
                    If lobjUltimoAno.ObjIdAnoShr.ObjValorPro > ObjAnoActual.ObjIdAnoShr.ObjValorPro Then
                        EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuParaCierreMes
                    Else
                        EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuCrearAno
                    End If
                Else
                    EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuParaCierreMes
                End If
            Case ClsOrionCop.FblnDebeCausarInt
                EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuCausarInt
            Case ObjAnoActual.FblnDebeCalcularCuotas
                EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuSinCalcAdmin
            Case ObjAnoActual.FblnDebeAjustarCuotasAdmin
                EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuDebeAjustarCuotasAdmin
            Case ObjAnoActual.FblnDebeImportarAjuste
                EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuDebeImportarAjuste
            Case ClsOrionCop.FblnHayItemsPorFacturar()
                EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuHayItemsProgFactPorProcesar
            Case ClsOrionCop.FblnCrearAno()
                EnuEstadoAplicacion = EnuEstadoAplicacionDef.EnuCrearAno
            Case Else
                astrMens = String.Empty
        End Select
        GobjPanDat.SControleProcesoObj(False)
    End Sub

    Private Function FblbImportarNotasDbIni() As Boolean
        Dim lblnImpo = FdblTasaMoraPeriodoActual() > 0 AndAlso (Not FblnHayNotasDb()) AndAlso
                FblnHayFacturas()
        lblnImpo = lblnImpo AndAlso ObjFechaUltCausacionGralDtm.ObjValorPro = GCDTMFECHANULA
        Return lblnImpo
    End Function

    Private Shared Function FblnCuentasContOk() As Boolean
        Dim lstrFiltro = ClsIdCarpetaCuentaShr.SstrNombreCampoBd & " = " & GshrIdCarpeta.ToString
        Dim ldrwCtasTerm() As DataRow = ClsPanorama.FdrwDataRow(ClsCuentaContabilidad.SstrNombreTabla,
                {ClsIdCuentaContStr.SstrNombreCampoBd}, {{ClsIdCuentaContStr.SstrNombreCampoBd, "ASC"}},
                lstrFiltro)
        Return (ldrwCtasTerm.Length >= 8)
    End Function

    Private Shared Function FblnCuentasBancosOk() As Boolean
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion
        Dim ldrwCtasBancos() As DataRow = ClsPanorama.FdrwDataRow(ClsCuentaBanco.SstrNombreTabla,
                {ClsIdCuentaBancoShr.SstrNombreCampoBd}, {{ClsIdCuentaBancoShr.SstrNombreCampoBd, "ASC"}},
                lstrFiltro)
        Return (ldrwCtasBancos.Length >= 1)
    End Function

    ''' <summary>
    ''' Verifica que al menos uno de los módulos de la copropiedad tengan  al menos un sector que le contribuya.
    ''' </summary>
    ''' <remarks></remarks>
    Private Function FblnSectoresModulos() As Boolean
        Dim lblnOk = True
        If IsNothing(McolModulos) Then
            McolModulos = ColModulos
        End If
        Dim lobjModulo As ClsModuloContribucion
        If McolModulos.Count > 0 Then
            For i = 1 To McolModulos.Count
                lobjModulo = McolModulos(i)
                If lobjModulo.ObjContribuyeCuotaAdminBln.ObjValorPro Then
                    If lobjModulo.ColSectoresModulo.Count = 0 Then
                        lblnOk = False
                        Exit For
                    End If
                End If
            Next
        End If
        Return lblnOk
    End Function

    Private Shared Function FblnTercerosCreados() As Boolean
        Dim ldrwTerceros() As DataRow = ClsPanorama.FdrwDataRow(ClsTercero.SstrNombreTabla,
                {ClsIdTerceroDbl.SstrNombreCampoBd}, {{ClsIdTerceroDbl.SstrNombreCampoBd, "ASC"}}, "")
        Return (ldrwTerceros.Length > 3)
    End Function

    Private Shared Function FblnClientesCreados() As Boolean
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion
        Dim ldrwClientes() As DataRow = ClsPanorama.FdrwDataRow(ClsCliente.SstrNombreTabla,
                {ClsIdClienteDbl.SstrNombreCampoBd}, {{ClsIdClienteDbl.SstrNombreCampoBd, "ASC"}},
                lstrFiltro)
        Return (ldrwClientes.Length > 1)
    End Function

    Private Shared Function FblnPrediosCreados() As Boolean
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion
        Dim ldrwPredios() As DataRow = ClsPanorama.FdrwDataRow(ClsPredio.SstrNombreTabla,
                {ClsIdPredioStr.SstrNombreCampoBd}, {{ClsIdPredioStr.SstrNombreCampoBd, "ASC"}},
                lstrFiltro)
        Return (ldrwPredios.Length > 0)
    End Function

    ''' <summary>
    ''' Indica si los coeficientes de propiedad de los predios esten calculados.
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared Function FblnCPCalculados() As Boolean
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion
        Dim ldtbCPPredios = ClsPanorama.FdtbDataTable(ClsPredio.SstrNombreTabla, {"SUM(" &
                ClsCoeficientePropiedadDec.SstrNombreCampoBd & ")"}, Nothing, lstrFiltro)
        Dim ldrwPredios() As DataRow = ldtbCPPredios.Select
        Dim ldecCoefPropTotal = Math.Round(ClsPanorama.FobjValorCampo(ldrwPredios(0)(0), EnuTipoValor.EnuDecimal), 4)
        Return (ldecCoefPropTotal = 1D)
    End Function

    Friend Shared Function FblnHayPropietarios() As Boolean
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion
        Dim lstrTabla = ClsPropietario.SstrNombreTabla
        Dim lstrExp = "SELECT COUNT(DISTINCT " & ClsIdPredioStr.SstrNombreCampoBd &
            ") FROM " & lstrTabla & " WHERE " & lstrFiltro
        Dim ldtbPredios = ClsPanorama.FdtbDataTable(lstrExp)
        Dim lentCantPreConProp As Integer = ClsPanorama.FobjValorCampo(ldtbPredios.Rows(0)(0),
                EnuTipoValor.EnuInteger)
        lstrTabla = ClsPredio.SstrNombreTabla
        ldtbPredios = ClsPanorama.FdtbDataTable(lstrTabla, {"COUNT(" &
                ClsIdPredio_PropStr.SstrNombreCampoBd & ")"}, {{"", ""}}, lstrFiltro)
        Dim lentCantPre As Integer = ClsPanorama.FobjValorCampo(ldtbPredios.Rows(0)(0),
                EnuTipoValor.EnuInteger)
        Dim lblnHay = lentCantPreConProp = lentCantPre
        Return lblnHay
    End Function

    Private Shared Function FblnHayTasaMora() As Boolean
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion
        Dim ldrwTasasMora() As DataRow = ClsPanorama.FdrwDataRow(ClsTasaMora.SstrNombreTabla,
                {ClsOrdinalTasaMoraEnt.SstrNombreCampoBd}, {{ClsOrdinalTasaMoraEnt.SstrNombreCampoBd, "ASC"}},
                lstrFiltro)
        Return (ldrwTasasMora.Length > 0)
    End Function

    Friend Shared Function FblnHayFacturas() As Boolean
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion
        Dim lblnHay As Boolean
        Dim lstrCamposSelect() = {"COUNT(" & ClsIdFacturaEnt.SstrNombreCampoBd & ")"}
        Dim ldtbFact As DataTable = ClsPanorama.FdtbDataTable(ClsFactura.SstrNombreTabla, lstrCamposSelect,
                Nothing, lstrFiltro)
        Dim ldrwRegs = ldtbFact.Select
        lblnHay = (ClsPanorama.FobjValorCampo(ldrwRegs(0)(0), EnuTipoValor.EnuInteger) > 0)
        Return lblnHay
    End Function

    Friend Shared Function FblnHayNotasDb() As Boolean
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion
        Dim lblnHay As Boolean
        Dim lstrCamposSelect() = {"COUNT(" & ClsIdNotaDbEnt.SstrNombreCampoBd & ")"}
        Dim ldtbFact As DataTable = ClsPanorama.FdtbDataTable(ClsNotaDb.SstrNombreTabla, lstrCamposSelect,
                Nothing, lstrFiltro)
        Dim ldrwRegs = ldtbFact.Select
        lblnHay = (ClsPanorama.FobjValorCampo(ldrwRegs(0)(0), EnuTipoValor.EnuInteger) > 0)
        Return lblnHay
    End Function

    Friend Function FblnGeneraInt() As Boolean
        Dim lblnGenInt = False
        For Each lobjServicio As ClsServicio In ObjAnoActual.ColServiciosAno
            If lobjServicio.FblnCausaMora Then
                lblnGenInt = True
                Exit For
            End If
        Next
        If Not lblnGenInt Then
            For Each lobjServicio As ClsServicio In ColServiciosPer
                If lobjServicio.FblnCausaMora Then
                    lblnGenInt = True
                    Exit For
                End If
            Next
        End If
        Return lblnGenInt
    End Function

    Friend Function FblnParametrosInterfazOk() As Boolean
        Dim lobjDoc As ClsDocumento, lblnOk = True
        For i = 1 To ColDocumentos.Count
            lobjDoc = GobjParametros.ColDocumentos(i)
            lblnOk = FblnPropiedadesDocOk(lobjDoc)
            If Not lblnOk Then
                Exit For
            End If
        Next
        Return lblnOk
    End Function

    Private Shared Function FblnPropiedadesDocOk(aobjDoc As ClsDocumento) As Boolean
        Dim lblnOk = True
        For Each lobjProp As ClsCBPropiedad In aobjDoc.ColPropiedades
            lblnOk = lobjProp.BlnEsValido
            If Not lblnOk Then
                Exit For
            End If
        Next
        Return lblnOk
    End Function

    Friend Function FblnServicioIdOk(ByRef astrMens As String) As Boolean
        Dim lblnSerIdOk = False
        GobjPanDat.SControleProcesoObj(True)
        If ObjServicioIdActivoBln.ObjValorPro Then
            If Not FblnHayServicioId() Then
                astrMens = "Es necesario crear el Servicio de Identificación!"
            ElseIf Not ClsOrionCop.FblnTodosPrediosAgruConValorId Then
                astrMens = "Es necesario asignar el valor del Servicio de Identificación " &
                        "a todos los Predios Agrupadores!"
            Else
                lblnSerIdOk = True
            End If
        Else
            lblnSerIdOk = True
        End If
        GobjPanDat.SControleProcesoObj(False)
        Return lblnSerIdOk
    End Function

    Private Function FblnServiciosOk(ByRef astrMens As String) As Boolean
        Dim lblnEsSerPermanente = False, lshrIdServicio = 0
        Dim lblnOk = True
        If Not IsNothing(ObjAnoActual) Then
            lblnOk = FblnSerBienParam(lblnEsSerPermanente, lshrIdServicio)
            If Not lblnOk Then
                astrMens = "El Servicio "
                If lblnEsSerPermanente Then
                    astrMens += "permanente " & lshrIdServicio.ToString
                Else
                    astrMens += lshrIdServicio.ToString & " del presente Año "
                End If
                astrMens += " tiene problemas de configuración. Es necesario corregirlo!"
            End If
        End If
        Return lblnOk
    End Function

    Private Function FblnSerBienParam(ByRef ablnEsper As Boolean, ByRef ashrIdSer As Short) As Boolean
        Dim lblnOk = False
        For Each lobjServicio As ClsServicio In ObjAnoActual.ColServiciosAno
            lblnOk = lobjServicio.FblnEstanTodosOk
            If Not lblnOk Then
                ablnEsper = False
                ashrIdSer = lobjServicio.ObjIdServicioShr.ObjValorPro
                Exit For
            End If
        Next
        If lblnOk Then
            For Each lobjServicio As ClsServicio In ColServiciosPer
                lblnOk = lobjServicio.FblnEstanTodosOk
                If Not lblnOk Then
                    ablnEsper = True
                    ashrIdSer = lobjServicio.ObjIdServicioShr.ObjValorPro
                    Exit For
                End If
            Next
        End If
        Return lblnOk
    End Function

    Private Function FblnSerBienParam() As Boolean
        Dim lblnOk = False
        If Not IsNothing(ObjAnoActual) Then
            If Not FblnPerActEsDicPrimerAno() Then
                For Each lobjServicio As ClsServicio In ObjAnoActual.ColServiciosAno
                    lblnOk = lobjServicio.FblnEstanTodosOk
                    If Not lblnOk Then
                        Exit For
                    End If
                Next
            Else
                lblnOk = True
            End If
        End If
        If lblnOk Then
            For Each lobjServicio As ClsServicio In ColServiciosPer
                lblnOk = lobjServicio.FblnEstanTodosOk
                If Not lblnOk Then
                    Exit For
                End If
            Next
        End If
        Return lblnOk
    End Function

    Private Function FblnServiciosPerPorCalcular(ByRef astrMens As String) As Boolean
        Dim lblnHaySerPorCal = False
        For Each lobjServicioPer As ClsServicio In ColServiciosPer
            If lobjServicioPer.ObjGeneraProgramBln.ObjValorPro Then
                lblnHaySerPorCal = Not lobjServicioPer.ObjEstaGenaradaProgramBln.ObjValorPro
                If lblnHaySerPorCal Then
                    astrMens = "Se deben calcular los valores a cobrar del servicio " & Chr(34) &
                            lobjServicioPer.ObjIdServicioShr.ToString & " " &
                            lobjServicioPer.ObjNombreServicioStr.ToString & Chr(34)
                    Exit For
                End If
            End If
        Next
        Return lblnHaySerPorCal
    End Function

    Friend Sub SVerifiqueResDian(ByRef astrMens As String)
        If GobjParametros.BlnEFacAutorizado Then
            Dim lenuEstadoResDian = GobjParametros.FEnuEstadoResDian
            If lenuEstadoResDian = EnuEstadoResDian.EnuSinResVigente Then
                astrMens = "Noy hay Resolución vigente de la DIAN!"
            ElseIf lenuEstadoResDian = EnuEstadoResDian.EnuVencida Then
                astrMens = "La Resolución de la DIAN está vencida!"
            ElseIf lenuEstadoResDian = EnuEstadoResDian.EnuNumPorAgotarse Then
                astrMens = "Según Resolución de la DIAN, hay menos de cuatrocientos Números " &
                        "disponibles!"
            ElseIf lenuEstadoResDian = EnuEstadoResDian.EnuNumAgotada Then
                astrMens = "Según Resolución de la DIAN, no hay Numeración disponible para Facturas!"
            End If
        End If
    End Sub

    Friend Function FEnuEstadoResDian() As EnuEstadoResDian
        Dim lenuEstadoResDian = EnuEstadoResDian.EnuOk
        If BlnEFacAutorizado Then
            If Not ClsOrionCop.FblnHayResVigente Then
                lenuEstadoResDian = EnuEstadoResDian.EnuSinResVigente
            Else
                Dim lentCantFacsDisponibles = ClsOrionCop.FentCantNumerosFacDisponibles
                If lentCantFacsDisponibles < 1 Then
                    lenuEstadoResDian = EnuEstadoResDian.EnuNumAgotada
                ElseIf lentCantFacsDisponibles < 400 Then
                    lenuEstadoResDian = EnuEstadoResDian.EnuNumPorAgotarse
                End If
                Dim lentDiasVigenciaResDian = ClsOrionCop.FentCantDiasVigenciaRes
                If lentDiasVigenciaResDian < 0 Then
                    lenuEstadoResDian = EnuEstadoResDian.EnuVencida
                ElseIf lentDiasVigenciaResDian = 0 Then
                    lenuEstadoResDian = EnuEstadoResDian.EnuVenceHoy
                ElseIf lentDiasVigenciaResDian <= 5 Then
                    lenuEstadoResDian = EnuEstadoResDian.EnuResPorVencer
                End If
            End If
        End If
        Return lenuEstadoResDian
    End Function
#End Region
End Class

#Region "Clases de Propiedad"
Friend Class ClsAutorizaEFacBln
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsCentroUtilOriCop = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "AutorizaEFac"
    Friend Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Autoriza Interfaz eFac"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If Not e.BlnVaciandoObjeto Then
            MobjPadre.ObjExigeFechaHoyDocsBln.SValide()
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return "Falso"
        Else
            Return ClsPanorama.FstrBuleanoToString(HobjValorPro)
        End If
    End Function
End Class

Friend Class ClsBaseRedondeoCPByt
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "BaseRedondeoCP"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = "BaseRedondeoCP"
        HblnRegistrarLogCambio = True
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, 9, BlnEsRequerido, HenuTipoValor)
        If ObjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If HblnEsValido Then
                Dim ldblRes As Double = HobjValorNew - Int(HobjValorNew)
                HblnEsValido = ldblRes = 0
            End If
            If Not HblnEsValido Then
                HstrMens = "El Valor de la Base de Redondeo debe ser un numero enero entre 1 y 9!"
                SNotifiqueDatInv()
            End If
        End If
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return ClsOrionCop.FstrNombreDatoConstanteOri(
                    EnuGrupoConstantesOriDef.EnuTipoBaseCalculo, HobjValorPro)
        End If
    End Function
End Class

Friend Class ClsBaseRedondeoGeneralDbl
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "BaseRedondeoGeneral"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = "BaseRedondeoGeneral"
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0.1, Integer.MaxValue, BlnEsRequerido,
                HenuTipoValor)
        If ObjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If HblnEsValido Then
                If HobjValorNew < 1 Then
                    Dim ldecValorPermitido = HobjValorNew * 10
                    HblnEsValido = False
                    For i = 0 To 2
                        If ldecValorPermitido = i Then
                            HblnEsValido = True
                            Exit For
                        End If
                    Next
                    If Not HblnEsValido Then
                        HstrMens = "Solo se permite redondear hasta 2 Decimales!"
                        SNotifiqueDatInv()
                    End If
                Else
                    HblnEsValido = (HobjValorNew = Int(HobjValorNew))
                    If Not HblnEsValido Then
                        HstrMens = "Si el Valor es mayor o igual a 1, este debe ser un Número entero!"
                        SNotifiqueDatInv()
                    End If
                End If
            End If
        End If
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsBaseRedondeoIntMoraDbl
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "BaseRedondeoInteresesMora"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = "BaseRedondeoIntMora"
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0.1, Integer.MaxValue, BlnEsRequerido,
                HenuTipoValor)
        If ObjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If HblnEsValido Then
                If HobjValorNew < 1 Then
                    Dim ldecValorPermitido = (HobjValorNew * 10)
                    HblnEsValido = False
                    For i = 0 To 2
                        If ldecValorPermitido = i Then
                            HblnEsValido = True
                            Exit For
                        End If
                    Next
                    If Not HblnEsValido Then
                        HstrMens = "Si el Valor es menor a 1 este debe ser 0, 0.1 o 0.2!"
                        SNotifiqueDatInv()
                    End If
                Else
                    HblnEsValido = (HobjValorNew = Int(HobjValorNew))
                    If Not HblnEsValido Then
                        HstrMens = "Si el Valor es mayor o igual a 1, este debe ser un Número entero!"
                        SNotifiqueDatInv()
                    End If
                End If
            End If
        End If
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

' Dato Requerido por ContaPyme
Friend Class ClsCodigoEmpShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "CodigoEmpresa"
    Private ReadOnly MobjPadre As ClsCentroUtilOriCop = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Codigo Empresa"
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HenuTipoValor = EnuTipoValor.enuShort
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsRequerido = (MobjPadre.ObjIdAppContableByt.ObjValorPro = EnuAppConta.EnuContaPyme)
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue,
                    BlnEsRequerido, EnuTipoValor)
        If Not BlnLeyendoOrigen Then
            If HblnEsValido Then
                HblnEsValido = (HobjValorNew - Int(HobjValorNew)) = 0
                If Not HblnEsValido Then
                    HstrMens = "El Código de la Empresa debe ser un Número entero!"
                    SNotifiqueDatInv()
                End If
            Else
                HstrMens = "El Código de la Empresa ingresado no es válido!"
                SNotifiqueDatInv()
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsConsolidaItemsFacBln
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Consolida Items Factura"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = "ConsolidaItemsFac"
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return "Falso"
        Else
            Return ClsPanorama.FstrBuleanoToString(HobjValorPro)
        End If
    End Function
End Class

Friend Class ClsDiasParaPersuasivoShr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "DiasParaPersuasivo"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = "DiasParaPersuasivo"
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue, BlnEsRequerido)
        If ObjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If HblnEsValido Then
                HblnEsValido = (HobjValorNew - Int(HobjValorNew) = 0)
                If Not HblnEsValido Then
                    HstrMens = "Para Persuasivo, el Dato ingresado debe ser un Número Entero!"
                    SNotifiqueDatInv()
                End If
            Else
                HstrMens = "Para Persuasivo, el Dato ingresado debe ser un Número entero mayor a uno!"
                SNotifiqueDatInv()
            End If
        End If
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsDiasParaPrejuridicoShr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "DiasParaPrejuridico"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = "DiasParaPrejuridico"
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsCentroUtilOriCop = ObjPadre
        Dim lshrCanDiasMin As Short = 1
        If lobjPadre.ObjDiasParaPersuasivoShr.BlnEsValido Then
            lshrCanDiasMin = lobjPadre.ObjDiasParaPersuasivoShr.ObjValorPro + 1
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, lshrCanDiasMin, Short.MaxValue,
                BlnEsRequerido)
        If ObjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If HblnEsValido Then
                HblnEsValido = (HobjValorNew - Int(HobjValorNew) = 0)
                If Not HblnEsValido Then
                    HstrMens = "Para Prejuridico, el Dato ingresado debe ser un Número Entero!"
                    SNotifiqueDatInv()
                End If
            Else
                If HobjValorNew <= lshrCanDiasMin Then
                    HstrMens = "Para Prejuridico, el Dato ingresado debe ser igual o mayor a " &
                            lshrCanDiasMin.ToString & "!"
                    SNotifiqueDatInv()
                End If
            End If
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If Not e.BlnVaciandoObjeto Then
            Dim lobjPadre As ClsCentroUtilOriCop = ObjPadre
            lobjPadre.ObjDiasParaPersuasivoShr.SValide()
        End If
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsDiasParaJuridicoShr
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsCentroUtilOriCop = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "DiasParaJuridico"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = "DiasParaJuridico"
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lshrCanDiasMin As Short = 1
        If MobjPadre.ObjDiasParaPrejuridicoShr.BlnEsValido Then
            lshrCanDiasMin = MobjPadre.ObjDiasParaPrejuridicoShr.ObjValorPro + 1
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, lshrCanDiasMin, Short.MaxValue,
                BlnEsRequerido)
        If ObjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If HblnEsValido Then
                HblnEsValido = (HobjValorNew - Int(HobjValorNew) = 0)
                If Not HblnEsValido Then
                    HstrMens = "Para Jurídico, el Dato ingresado debe ser un Número Entero!"
                    SNotifiqueDatInv()
                End If
            Else
                If HobjValorNew <= lshrCanDiasMin Then
                    HstrMens = "Para Jurídico, el Dato ingresado debe ser igual o mayor a " &
                            lshrCanDiasMin.ToString & "!"
                    SNotifiqueDatInv()
                End If
            End If
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If Not e.BlnVaciandoObjeto AndAlso HblnEsValido Then
            Dim lobjPadre As ClsCentroUtilOriCop = ObjPadre
            lobjPadre.ObjDiasParaPrejuridicoShr.SValide()
        End If
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsDiasParaPerdidaShr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "DiasParaPerdida"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = "DiasParaPerdida"
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsCentroUtilOriCop = ObjPadre
        Dim lshrCanDiasMin As Short = 1
        If lobjPadre.ObjDiasParaJuridicoShr.BlnEsValido Then
            lshrCanDiasMin = lobjPadre.ObjDiasParaJuridicoShr.ObjValorPro + 1
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, lshrCanDiasMin, Short.MaxValue,
                BlnEsRequerido)
        If ObjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If HblnEsValido Then
                HblnEsValido = (HobjValorNew - Int(HobjValorNew) = 0)
                If Not HblnEsValido Then
                    HstrMens = "Para Perdida, el Dato ingresado debe ser un Número Entero!"
                    SNotifiqueDatInv()
                End If
            Else
                If HobjValorNew <= lshrCanDiasMin Then
                    HstrMens = "Para Perdida, el Dato ingresado debe ser igual o mayor a " &
                            lshrCanDiasMin.ToString & "!"
                    SNotifiqueDatInv()
                End If
            End If
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If Not e.BlnVaciandoObjeto AndAlso HblnEsValido Then
            Dim lobjPadre As ClsCentroUtilOriCop = ObjPadre
            lobjPadre.ObjDiasParaJuridicoShr.SValide()
        End If
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsExigeFechaHoyCajaBln
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Exige Fecha Hoy Caja"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = "ExigeFechaHoyCaja"
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return "Falso"
        Else
            Return ClsPanorama.FstrBuleanoToString(HobjValorPro)
        End If
    End Function
End Class

Friend Class ClsExigeFechaHoyDocsBln
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsCentroUtilOriCop = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Exige Fecha Hoy en Documentos"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = "ExigeFechaHoyDocs"
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If Not HobjValorNew Then
                HblnEsValido = Not MobjPadre.BlnEFacAutorizado
                If Not HblnEsValido Then
                    HstrMens = "Cuando está habilitada la Facturación Electrónica es necesario " &
                            "que los Documentos sean expedidos con la Fecha de Hoy!"
                    SNotifiqueDatInv()
                End If
            End If
            If HblnEsValido Then
                If (Not HobjValorOriginal) AndAlso (HobjValorNew) Then
                    If MobjPadre.ObjAnoActual IsNot Nothing Then
                        Dim lstrPerAct = MobjPadre.ObjAnoActual.StrIdPeriodoActual
                        Dim lstrPerHoy = ClsPanorama.FstrPeriodo(Date.Today)
                        HblnEsValido = (lstrPerAct = lstrPerHoy)
                        If Not HblnEsValido Then
                            HstrMens = "Es necesario que el Periodo Actual del Programa " &
                                    "corresponda con el Mes de la Fecha de Hoy"
                            SNotifiqueDatInv()
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return "Falso"
        Else
            Return ClsPanorama.FstrBuleanoToString(HobjValorPro)
        End If
    End Function
End Class

Friend Class ClsFechaResolucionContDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaResolContingencia"
    Private ReadOnly MobjPadre As ClsCentroUtilOriCop = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Fecha Resolución Contingencia"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MyBase.SVaciePropiedad()
        HobjValorNew = GCDTMFECHANULA
        HobjValorPro = GCDTMFECHANULA
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = True
        HblnEsRequerido = MobjPadre.BlnEFacAutorizado
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            Dim ldtmFechaMin As Date = DateSerial(2017, 12, 31)
            Dim ldtmFechaMax As Date = Date.Today
            HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax,
                    BlnEsRequerido)
            If Not HblnEsValido Then
                HstrMens = "La Fecha ingresada no es válida!"
                SNotifiqueDatInv()
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(HobjValorPro) Then
            Return Format(HobjValorPro, "yyyy-MM-dd")
        Else
            Return Format(GCDTMFECHANULA, GCSTRFMTFECHASIMPLE)
        End If
    End Function
End Class

Friend Class ClsFechaResolucionFactDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaResolFactura"
    Private ReadOnly MobjPadre As ClsCentroUtilOriCop = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Fecha Resolución Facturación"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MyBase.SVaciePropiedad()
        HobjValorNew = GCDTMFECHANULA
        HobjValorPro = GCDTMFECHANULA
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = True
        HblnEsRequerido = Not String.IsNullOrEmpty(MobjPadre.ObjNumeroResolFacturaStr.ToString)
        HstrMens = String.Empty
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            Dim ldtmFechaMin As Date = DateSerial(2019, 12, 31)
            Dim ldtmFechaMax As Date = Date.Today
            HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax,
                    BlnEsRequerido)
            If Not HblnEsValido Then
                HstrMens = "La Fecha ingresada no es válida!"
            End If
        ElseIf MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
            If MobjPadre.BlnEFacAutorizado AndAlso HobjValorNew <> HobjValorOriginal Then
                HblnEsValido = HobjValorNew > HobjValorOriginal
            End If
            If Not HblnEsValido Then
                HstrMens = "La Fecha ingresada no es válida!"
            End If
        End If
        If Not String.IsNullOrEmpty(HstrMens) Then
            SNotifiqueDatInv()
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(HobjValorPro) Then
            Return Format(HobjValorPro, "yyyy-MM-dd")
        Else
            Return Format(GCDTMFECHANULA, GCSTRFMTFECHASIMPLE)
        End If
    End Function
End Class

Friend Class ClsFechaVenceResolFactDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaVenceResolFact"
    Private ReadOnly MobjPadre As ClsCentroUtilOriCop = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Fecha Vence Resolución Facturación"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MyBase.SVaciePropiedad()
        HobjValorNew = GCDTMFECHANULA
        HobjValorPro = GCDTMFECHANULA
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = True
        HblnEsRequerido = Not String.IsNullOrEmpty(MobjPadre.ObjNumeroResolFacturaStr.ToString)
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            Dim ldtmFechaMin As Date = GCDTMFECHANULA
            Dim ldtmFechaMax As Date = Date.Today.AddYears(2)
            HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax,
                    BlnEsRequerido)
            If HblnEsValido Then
                If HblnEsRequerido AndAlso MobjPadre.ObjFechaResolucionFactDtm.BlnEsValido Then
                    HblnEsValido = HobjValorNew > MobjPadre.ObjFechaResolucionFactDtm.ObjValorPro
                    If Not HblnEsValido Then
                        HstrMens = "La Fecha ingresada no es válida!"
                        SNotifiqueDatInv()
                    End If
                End If
            Else
                HstrMens = "La Fecha ingresada no es válida!"
                SNotifiqueDatInv()
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(HobjValorPro) Then
            Return Format(HobjValorPro, "yyyy-MM-dd")
        Else
            Return Format(GCDTMFECHANULA, GCSTRFMTFECHASIMPLE)
        End If
    End Function
End Class

Friend Class ClsFechaUltCausacionGralDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaUltimaCausacionGral"
    Private ReadOnly MobjPadre As ClsCentroUtilOriCop = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Fecha última causación general"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MyBase.SVaciePropiedad()
        HobjValorNew = GCDTMFECHANULA
        HobjValorPro = GCDTMFECHANULA
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = True
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            Dim ldtmFechaMin As Date = DateSerial(2019, 12, 31)
            Dim ldtmFechaMax As Date = Date.Today
            HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax,
                    BlnEsRequerido)
            If Not HblnEsValido Then
                HstrMens = "La Fecha ingresada no es válida!"
                SLevanteEveNot("", 0, EnuSeveridadNot.EnuError)
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(HobjValorPro) Then
            Return Format(HobjValorPro, GCSTRFMTFECHASIMPLE)
        Else
            Return Format(GCDTMFECHANULA, GCSTRFMTFECHASIMPLE)
        End If
    End Function
End Class

Friend Class ClsFirmaRCeMail
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Firma Recibo de Caja Correo"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = "FirmaRCEMail"
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
        If HblnEsValido AndAlso HobjValorNew Then
            HblnEsValido = ClsPanorama.FblnEmailsHabilitado
            If Not HblnEsValido Then
                HstrMens = "No está habilitado para enviar Documentos por Email!"
                SNotifiqueDatInv()
            End If
        End If
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return "Falso"
        Else
            Return ClsPanorama.FstrBuleanoToString(HobjValorPro)
        End If
    End Function
End Class

Friend Class ClsIdAppContableByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdAppContable"
    Private ReadOnly MobjPadre As ClsCentroUtilOriCop = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Id App Contable"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsRequerido = MobjPadre.ObjTipoInterfazByt.ObjValorPro > EnuTipoInterfazDef.None
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew,
                EnuAppConta.EnuApoloAP, EnuAppConta.EnuPodium, BlnEsRequerido)
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If Not e.BlnVaciandoObjeto AndAlso HblnEsValido Then
            MobjPadre.ObjTipoInterfazByt.SValide()
        End If
    End Sub
    Private Sub ClsIdAppContableByt_evnPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If Not e.BlnVaciandoObjeto Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                MobjPadre.ObjCodigoEmpShr.SValide()
                MobjPadre.ObjTipoInterfazByt.SValide()
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Friend ReadOnly Property StrNombreApp As String
        Get
            Dim lstrNombreApp = String.Empty
            If HblnEsValido Then
                lstrNombreApp = ClsOrionCop.FstrNombreDatoConstanteOri(EnuGrupoConstantesOriDef.EnuAppContable,
                        HobjValorPro)
            End If
            Return lstrNombreApp
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsIdCtaAnticiposRecibidosStr
    Inherits ClsCBPropiedad
    Private MstrNombreCuenta As String = String.Empty
    Private Const MCSTRNOMBRECAMPOBD As String = "IdCtaAnticipos"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "CuentaAnticiposRec"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        MstrNombreCuenta = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 4, ShrLongitud, BlnEsRequerido)
        If ObjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If ObjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                If HblnEsValido Then
                    HblnEsValido = ClsOrionCop.FblnEsValidaCtaContabilidad(HobjValorNew.ToString)
                    If HblnEsValido Then
                        MstrNombreCuenta = ClsOrionCop.FstrNombreCuentaCon(HobjValorNew.ToString)
                    Else
                        HstrMens = "La Cuenta de Contabilidad ingresada no ha sido creada aún!"
                        SNotifiqueDatInv()
                    End If
                Else
                    HstrMens = "El Dato ingresado no es válido. Debe ser una Cadena de 4 a 30 Caracteres!"
                    SNotifiqueDatInv()
                End If
            End If
        Else
            If HblnEsValido Then
                MstrNombreCuenta = ClsOrionCop.FstrNombreCuentaCon(HobjValorNew.ToString)
            End If
        End If
    End Sub
    Friend ReadOnly Property StrNombreCuenta As String
        Get
            Return MstrNombreCuenta
        End Get
    End Property
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsIdCtaCajaStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdCtaCaja"
    Private MstrNombreCuenta As String = String.Empty
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "CuentaCaja"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        MstrNombreCuenta = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 4, ShrLongitud, BlnEsRequerido)
        If ObjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If ObjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                If HblnEsValido Then
                    HblnEsValido = ClsOrionCop.FblnEsValidaCtaContabilidad(HobjValorNew.ToString)
                    If HblnEsValido Then
                        MstrNombreCuenta = ClsOrionCop.FstrNombreCuentaCon(HobjValorNew.ToString)
                    Else
                        HstrMens = "La Cuenta de Contabilidad ingresada no ha sido creada aún!"
                        SNotifiqueDatInv()
                    End If
                Else
                    HstrMens = "El Dato ingresado no es válido. Debe ser una Cadena de 4 a 30 Caracteres!"
                    SNotifiqueDatInv()
                End If
            End If
        Else
            If HblnEsValido Then
                MstrNombreCuenta = ClsOrionCop.FstrNombreCuentaCon(HobjValorNew.ToString)
            End If
        End If
    End Sub
    Friend ReadOnly Property StrNombreCuenta As String
        Get
            Return MstrNombreCuenta
        End Get
    End Property
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsIdCtaDescuentosPPStr
    Inherits ClsCBPropiedad
    Private MstrNombreCuenta As String = String.Empty
    Private Const MCSTRNOMBRECAMPOBD As String = "IdCtaDctoPP"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Cuenta Descuentos Pronto Pago"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        MstrNombreCuenta = String.Empty
        Dim lobjPadre As ClsCentroUtilOriCop = ObjPadre
        If Not IsNothing(lobjPadre.ObjAnoActual) Then
            HblnEsRequerido = GobjParametros.ObjAnoActual.ObjAplicaDsctoPPBln.ObjValorPro
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 4, ShrLongitud, BlnEsRequerido)
        If ObjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If HblnEsValido Then
                If BlnEsRequerido OrElse Not String.IsNullOrEmpty(HobjValorNew) Then
                    HblnEsValido = ClsOrionCop.FblnEsValidaCtaContabilidad(HobjValorNew.ToString)
                    If HblnEsValido Then
                        MstrNombreCuenta = ClsOrionCop.FstrNombreCuentaCon(HobjValorNew.ToString)
                    Else
                        HstrMens = "La Cuenta de Contabilidad ingresada no ha sido creada aún!"
                        SNotifiqueDatInv()
                    End If
                End If
            Else
                HstrMens = "El Dato ingresado no es válido. Debe ser una Cadena de 4 a 30 Caracteres!"
                SNotifiqueDatInv()
            End If
        Else
            If HblnEsValido Then
                MstrNombreCuenta = ClsOrionCop.FstrNombreCuentaCon(HobjValorNew.ToString)
            End If
        End If
    End Sub
    Friend ReadOnly Property StrNombreCuenta As String
        Get
            Return MstrNombreCuenta
        End Get
    End Property
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsIdCtaImptosAsumidosStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IDCtaImpAsumidos"
    Private MstrNombreCuenta As String = String.Empty
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Cta Impuestos Asumidos"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        MstrNombreCuenta = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 4, ShrLongitud, BlnEsRequerido)
        If ObjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If HblnEsValido Then
                If Not String.IsNullOrEmpty(HobjValorNew) Then
                    HblnEsValido = ClsOrionCop.FblnEsValidaCtaContabilidad(HobjValorNew.ToString)
                    If HblnEsValido Then
                        MstrNombreCuenta = ClsOrionCop.FstrNombreCuentaCon(HobjValorNew.ToString)
                    Else
                        HstrMens = "La Cuenta de Contabilidad ingresada no ha sido creada aún!"
                        SNotifiqueDatInv()
                    End If
                End If
            Else
                HstrMens = "El Dato ingresado no es válido. Debe ser una Cadena de 4 a 30 Caracteres!"
                SNotifiqueDatInv()
            End If
        Else
            If HblnEsValido Then
                MstrNombreCuenta = ClsOrionCop.FstrNombreCuentaCon(HobjValorNew.ToString)
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Friend ReadOnly Property StrNombreCuenta As String
        Get
            Return MstrNombreCuenta
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsIdCtaIngPorIdentificarStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IDCtaIngPorIdentificar"
    Private MstrNombreCuenta As String = String.Empty
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Cta Ingresos por Identificar"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        MstrNombreCuenta = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 4, ShrLongitud, BlnEsRequerido)
        If ObjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If HblnEsValido Then
                If Not String.IsNullOrEmpty(HobjValorNew) Then
                    HblnEsValido = ClsOrionCop.FblnEsValidaCtaContabilidad(HobjValorNew.ToString)
                    If HblnEsValido Then
                        MstrNombreCuenta = ClsOrionCop.FstrNombreCuentaCon(HobjValorNew.ToString)
                    Else
                        HstrMens = "La Cuenta de Contabilidad ingresada no ha sido creada aún!"
                        SNotifiqueDatInv()
                    End If
                End If
            Else
                HstrMens = "El Dato ingresado no es válido. Debe ser una Cadena de 4 a 30 Caracteres!"
                SNotifiqueDatInv()
            End If
        Else
            If HblnEsValido Then
                MstrNombreCuenta = ClsOrionCop.FstrNombreCuentaCon(HobjValorNew.ToString)
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Friend ReadOnly Property StrNombreCuenta As String
        Get
            Return MstrNombreCuenta
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsIdCtaIntMoraDbStr
    Inherits ClsCBPropiedad
    Private MstrNombreCuenta As String = String.Empty
    Private Const MCSTRNOMBRECAMPOBD As String = "IdCtaMoraDb"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "CuentaIntMoraDb"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        MstrNombreCuenta = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 4, ShrLongitud, BlnEsRequerido)
        If ObjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If HblnEsValido Then
                HblnEsValido = ClsOrionCop.FblnEsValidaCtaContabilidad(HobjValorNew.ToString)
                If HblnEsValido Then
                    MstrNombreCuenta = ClsOrionCop.FstrNombreCuentaCon(HobjValorNew.ToString)
                Else
                    HstrMens = "La Cuenta de Contabilidad ingresada no ha sido creada aún!"
                    SNotifiqueDatInv()
                End If
            Else
                HstrMens = "El Dato ingresado no es válido. Debe ser una Cadena de 4 a 30 Caracteres!"
                SNotifiqueDatInv()
            End If
        Else
            If HblnEsValido Then
                MstrNombreCuenta = ClsOrionCop.FstrNombreCuentaCon(HobjValorNew.ToString)
            End If
        End If
    End Sub
    Friend ReadOnly Property StrNombreCuenta As String
        Get
            Return MstrNombreCuenta
        End Get
    End Property
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsIdCtaReteFuenteStr
    Inherits ClsCBPropiedad
    Private MstrNombreCuenta As String = String.Empty
    Private Const MCSTRNOMBRECAMPOBD As String = "IdCtaRetefuente"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "CuentaReteFuente"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        MstrNombreCuenta = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 4, ShrLongitud, BlnEsRequerido)
        If ObjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If HblnEsValido Then
                If Not String.IsNullOrEmpty(HobjValorNew) Then
                    HblnEsValido = ClsOrionCop.FblnEsValidaCtaContabilidad(HobjValorNew.ToString)
                    If HblnEsValido Then
                        MstrNombreCuenta = ClsOrionCop.FstrNombreCuentaCon(HobjValorNew.ToString)
                    Else
                        HstrMens = "La Cuenta de Contabilidad ingresada no ha sido creada aún!"
                        SNotifiqueDatInv()
                    End If
                End If
            Else
                HstrMens = "El Dato ingresado no es válido. Debe ser una Cadena de 4 a 30 Caracteres!"
                SNotifiqueDatInv()
            End If
        Else
            If HblnEsValido Then
                MstrNombreCuenta = ClsOrionCop.FstrNombreCuentaCon(HobjValorNew.ToString)
            End If
        End If
    End Sub
    Friend ReadOnly Property StrNombreCuenta As String
        Get
            Return MstrNombreCuenta
        End Get
    End Property
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsIdCtaReteIcaStr
    Inherits ClsCBPropiedad
    Private MstrNombreCuenta As String = String.Empty
    Private Const MCSTRNOMBRECAMPOBD As String = "IdCtaReteIca"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "CuentaReteIca"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        MstrNombreCuenta = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 4, ShrLongitud, BlnEsRequerido)
        If ObjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If HblnEsValido Then
                If Not String.IsNullOrEmpty(HobjValorNew) Then
                    HblnEsValido = ClsOrionCop.FblnEsValidaCtaContabilidad(HobjValorNew.ToString)
                    If HblnEsValido Then
                        MstrNombreCuenta = ClsOrionCop.FstrNombreCuentaCon(HobjValorNew.ToString)
                    Else
                        HstrMens = "La Cuenta de Contabilidad ingresada no ha sido creada aún!"
                        SNotifiqueDatInv()
                    End If
                End If
            Else
                HstrMens = "El Dato ingresado no es válido. Debe ser una Cadena de 4 a 30 Caracteres!"
                SNotifiqueDatInv()
            End If
        Else
            If HblnEsValido Then
                MstrNombreCuenta = ClsOrionCop.FstrNombreCuentaCon(HobjValorNew.ToString)
            End If
        End If
    End Sub
    Friend ReadOnly Property StrNombreCuenta As String
        Get
            Return MstrNombreCuenta
        End Get
    End Property
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsIdCtaReteIvaStr
    Inherits ClsCBPropiedad
    Private MstrNombreCuenta As String = String.Empty
    Private Const MCSTRNOMBRECAMPOBD As String = "IdCtaReteIva"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "CuentaReteIva"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        MstrNombreCuenta = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 4, ShrLongitud, BlnEsRequerido)
        If ObjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If HblnEsValido Then
                If Not String.IsNullOrEmpty(HobjValorNew) Then
                    HblnEsValido = ClsOrionCop.FblnEsValidaCtaContabilidad(HobjValorNew.ToString)
                    If HblnEsValido Then
                        MstrNombreCuenta = ClsOrionCop.FstrNombreCuentaCon(HobjValorNew.ToString)
                    Else
                        HstrMens = "La Cuenta de Contabilidad ingresada no ha sido creada aún!"
                        SNotifiqueDatInv()
                    End If
                End If
            Else
                HstrMens = "El Dato ingresado no es válido. Debe ser una Cadena de 4 a 30 Caracteres!"
                SNotifiqueDatInv()
            End If
        Else
            If HblnEsValido Then
                MstrNombreCuenta = ClsOrionCop.FstrNombreCuentaCon(HobjValorNew.ToString)
            End If
        End If
    End Sub
    Friend ReadOnly Property StrNombreCuenta As String
        Get
            Return MstrNombreCuenta
        End Get
    End Property
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsIdMedioPagoDefectoByt
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsCentroUtilOriCop = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "TipoMedioPagoDefecto"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Tipo Medio Pago por Defecto"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsRequerido = MobjPadre.ObjIdProvEFacByt.ObjValorPro > 0 AndAlso
                MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew,
                EnuTipoMedioPagoDef.EnuEfectivo, EnuTipoMedioPagoDef.EnuTransferencia, HblnEsRequerido)
        If Not HblnEsValido Then
            HstrMens = "El Medio de Pago ingresado no es válido!"
            SNotifiqueDatInv()
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(ObjValorPro) Then
            Return ClsOrionCop.FstrNombreDatoConstanteOri(
                    EnuGrupoConstantesOriDef.EnuMediosPago, HobjValorPro)
        Else
            Return ""
        End If
    End Function
End Class

Friend Class ClsIdProvEFacByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdProvEFac"
    Private ReadOnly MobjPadre As ClsCentroUtilOriCop = Nothing
    Private MobjProvEFactura As ClsProveedorEFac = Nothing
    Private MblnExisteProvEFac As Boolean = False
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Proveedor eFac"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew,
                EnuProveedorEFac.EnuProtecdataMisFac, EnuProveedorEFac.EnuProtecdataMisFac,
                BlnEsRequerido)
        If HblnEsValido Then
            If Not BlnLeyendoOrigen Then
                If HobjValorNew > 0 Then
                    If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                        If MobjProvEFactura.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                            SAbraProveedorEFac()
                            If Not MblnExisteProvEFac Then
                                MobjProvEFactura.SCreeObj({GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew})
                            Else
                                MobjProvEFactura.SModifique()
                            End If
                        End If
                    Else
                        SAbraProveedorEFac()
                    End If
                End If
            Else
                If HobjValorNew > 0 Then
                    SAbraProveedorEFac()
                End If
            End If
        Else
            HstrMens = "El Dato ingresado no es válido!"
            SNotifiqueDatInv()
        End If
    End Sub
    Friend ReadOnly Property ObjProvEFac As ClsProveedorEFac
        Get
            SAbraProveedorEFac()
            Return MobjProvEFactura
        End Get
    End Property
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        MobjPadre.ObjNumeroResolFacturaStr.SValide()
        MobjPadre.ObjFechaResolucionFactDtm.SValide()
        MobjPadre.ObjFechaVenceResolFactDtm.SValide()
        MobjPadre.ObjRangoFraIniEnt.SValide()
        MobjPadre.ObjRangoFraFinEnt.SValide()
        MobjPadre.ObjExigeFechaHoyDocsBln.SValide()
    End Sub
    Private Sub SAbraProveedorEFac()
        Dim lobjvalorLlave() As Object = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
        If IsNothing(MobjProvEFactura) Then
            MobjProvEFactura = New ClsProveedorEFac(ObjPadre, EnuModoInstanciaObjDef.enuUnico)
        End If
        If HblnEsValido Then
            If MobjProvEFactura.ObjIdProveedorEFacEnt.ObjValorPro <> HobjValorNew Then
                MobjProvEFactura.SAbra(lobjvalorLlave)
            End If
        End If
        MblnExisteProvEFac = MobjProvEFactura.BlnExiste
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsInformaSaldoTotalDespuesRCBln
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "InformaDeudaTotalDespuesRC"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = "InformaDeudaTotDespRC"
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return "Falso"
        Else
            Return ClsPanorama.FstrBuleanoToString(HobjValorPro)
        End If
    End Function
End Class

Friend Class ClsNoMostrarAyudaBln
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "No mostrar ayuda"
        HenuTipoValor = EnuTipoValor.EnuBoolean
        HStrNombreCampoBd = "NoMostrarAyuda"
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return "Falso"
        Else
            Return ClsPanorama.FstrBuleanoToString(HobjValorPro)
        End If
    End Function
End Class

Friend Class ClsNotificacionesSonorasBln
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Notificaciones Sonoras"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = "NotificacionesSonoras"
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return "Falso"
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsNumeroResolContiStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "NumeroResolContingencia"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Numero resolucion Contingencia"
        HshrLongitud = 20
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 4, ShrLongitud, BlnEsRequerido)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsNumeroResolFacturaStr
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsCentroUtilOriCop = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "NumeroResolFactura"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Numero resolucion Fras"
        HshrLongitud = 20
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsRequerido = (MobjPadre.BlnEFacAutorizado)
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 4, ShrLongitud, BlnEsRequerido)
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If Not HblnEsValido Then
                HstrMens = "Se debe introducir el Número de Resolución que autoriza la Numeración " &
                    "de las Facturas!"
                SNotifiqueDatInv()
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsPieFacturaDos_CUStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PieFacturaDos"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "PieFacturaDos"
        HshrLongitud = 230
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 3, ShrLongitud, BlnEsRequerido)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return String.Empty
        Else
            Return HobjValorPro.ToString().Trim
        End If
    End Function
End Class

Friend Class ClsPieFacturaUno_CUStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PieFacturaUno"
    Private ReadOnly MobjPadre As ClsCentroUtilOriCop = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "PieFacturaUno"
        HshrLongitud = 230
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 3, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando AndAlso
                    GobjParametros.FstrPieFacturaRes <> "" Then
                HblnEsValido = MobjPadre.FstrPieFacturaRes = HobjValorNew
                If Not HblnEsValido Then
                    HstrMens = "El Pie de Factura uno no puede ser " &
                            "modificado porque está construido a partir " &
                            "Resolución DIAN!"
                    SLevanteEveNot("", 0, EnuSeveridadNot.EnuDatoInvalido)
                End If
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString().Trim()
        End If
    End Function
End Class

Friend Class ClsParametrizacionOkBln
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "ParametrizacionOk"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = "ParametrizacionOk"
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return "Falso"
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsPermiteAnticipoPorServicioBln
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Permite Anticipos Por Servicio"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = "PermiteAnticipoPorSer"
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return "Falso"
        Else
            Return ClsPanorama.FstrBuleanoToString(HobjValorPro)
        End If
    End Function
End Class

Friend Class ClsPlazoDefectoFacManualShr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "PlazoDefectoFacturaMan"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = "PlazoDefectoFacManual"
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, 180, BlnEsRequerido)
        If HblnEsValido Then
            HblnEsValido = (HobjValorNew - Int(HobjValorNew) = 0)
        End If
        If Not HblnEsValido Then
            HstrMens = "El Dato debe ser un Número entero entre cero y ciento ochenta!"
            SNotifiqueDatInv()
        End If
        If Not HblnEsValido Then
        End If
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsPrefijoFactContStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PrefijoFactCont"
    Private ReadOnly MobjPadre As ClsCentroUtilOriCop = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Prefijo Factura Contingencia"
        HshrLongitud = 10
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = String.Empty
        HblnEsRequerido = Not String.IsNullOrEmpty(MobjPadre.ObjNumeroResolContiStr.ObjValorPro)
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud, BlnEsRequerido)
        If Not HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                If HobjValorNew.ToString > 10 Then
                    HstrMens = "El Prefijo del Documento debe tener una Longitud entre " &
                            "0 y 10 Caracteres!"
                    SNotifiqueDatInv()
                End If
            End If
        End If
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsRangoFraConFinEnt
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsCentroUtilOriCop = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "RangoFacturaConFin"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Rango Factura Contingencia Final"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsRequerido = (MobjPadre.ObjRangoFraConIniEnt.ObjValorPro > 0)
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue,
                BlnEsRequerido, EnuTipoValor.enuInteger)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                If BlnEsRequerido Then
                    HblnEsValido = (HobjValorNew > MobjPadre.ObjRangoFraConIniEnt.ObjValorPro)
                End If
                If HblnEsValido Then
                    HblnEsValido = (HobjValorNew - Int(HobjValorNew) = 0)
                    If Not HblnEsValido Then
                        HstrMens = "El Valor debe ser un Numero entero!"
                        SNotifiqueDatInv()
                    End If
                Else
                    HstrMens = "El Valor debe ser mayor al Número inicial!"
                    SNotifiqueDatInv()
                End If
            End If
        Else
            HstrMens = "EL Valor debe ser un Número entero mayor a cero!"
            SNotifiqueDatInv()
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsRangoFraConIniEnt
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsCentroUtilOriCop = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "RangoFacturaConIni"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Rango Factura Contingencia Inicial"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue,
                BlnEsRequerido, EnuTipoValor.enuInteger)
        If HblnEsValido Then
            HblnEsValido = (HobjValorNew - Int(HobjValorNew) = 0)
            If Not HblnEsValido Then
                HstrMens = "El Valor debe ser un Número entero!"
                SNotifiqueDatInv()
            End If
        Else
            HstrMens = "El Valor debe ser un Número Entero mayor a cero!"
            SNotifiqueDatInv()
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
            MobjPadre.ObjRangoFraConFinEnt.SValide()
        End If
    End Sub
End Class

Friend Class ClsRangoFraFinEnt
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsCentroUtilOriCop = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "RangoFacturaFin"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Rango Factura Final"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsRequerido = Not String.IsNullOrEmpty(MobjPadre.ObjNumeroResolFacturaStr.ToString)
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue,
                BlnEsRequerido, EnuTipoValor.enuInteger)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando AndAlso HblnEsValido Then
                If BlnEsRequerido Then
                    HblnEsValido = (HobjValorNew > MobjPadre.ObjRangoFraIniEnt.ObjValorPro)
                End If
                If HblnEsValido Then
                    HblnEsValido = (HobjValorNew - Int(HobjValorNew) = 0)
                    If Not HblnEsValido Then
                        HstrMens = "El Valor debe ser un Número entero mayor!"
                        SNotifiqueDatInv()
                    End If
                Else
                    HstrMens = "El valor debe ser mayor al Número inicial!"
                    SNotifiqueDatInv()
                End If
            End If
        Else
            HstrMens = "El Valor debe ser un Número entero mayor a cero!"
            SNotifiqueDatInv()
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsRangoFraIniEnt
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsCentroUtilOriCop = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "RangoFacturaIni"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Rango Factura Inicial"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsRequerido = Not String.IsNullOrEmpty(MobjPadre.ObjNumeroResolFacturaStr.ToString)
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue,
                BlnEsRequerido, EnuTipoValor.enuInteger)
        If HblnEsValido Then
            HblnEsValido = (HobjValorNew - Int(HobjValorNew) = 0)
            If Not HblnEsValido Then
                HstrMens = "El Valor debe ser un Número entero!"
                SNotifiqueDatInv()
            End If
        Else
            HstrMens = "El Valor debe ser un Número entero mayor a cero!"
            SNotifiqueDatInv()
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
            MobjPadre.ObjRangoFraFinEnt.SValide()
        End If
    End Sub
End Class

Friend Class ClsServicioIdActivoBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "ServicioIdActivo"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Esta Servicio Identificación activo"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return "Falso"
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsTarifaReteIvaDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "TarifaReteIva"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Tarifa ReteIva"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsCentroUtilOriCop = ObjPadre
        HblnEsRequerido = lobjPadre.BlnEFacAutorizado
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0,
                0.5, BlnEsRequerido, HenuTipoValor)
        If Not HblnEsValido Then
            HstrMens = "El Valor debe ser un Número decimal menor a punto cincuenta!"
            SNotifiqueDatInv()
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return Format(HobjValorPro, "p")
        End If
    End Function
End Class

Friend Class ClsTipoInterfazByt
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsCentroUtilOriCop = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Tipo Interfaz"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = "TipoInterfaz"
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew,
                EnuTipoInterfazDef.EnuPorComprobante, EnuTipoInterfazDef.EnuPorDocumento,
                BlnEsRequerido)
        HblnEsRequerido = MobjPadre.ObjIdAppContableByt.ObjValorPro > 0
        Dim lenuTipoInt As EnuTipoInterfazDef = EnuTipoInterfazDef.None
        If BlnEsRequerido Then
            lenuTipoInt = EnuTipoInterfazDef.EnuPorComprobante
        End If
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                HblnEsValido = HobjValorNew >= lenuTipoInt AndAlso
                                HobjValorNew <= EnuTipoInterfazDef.EnuPorDocumento
                If HblnEsValido Then
                    If MobjPadre.ObjIdAppContableByt.ObjValorPro = EnuAppConta.EnuContaPyme OrElse
                            MobjPadre.ObjIdAppContableByt.ObjValorPro = EnuAppConta.EnuSIIGO Then
                        HblnEsValido = (HobjValorNew = EnuTipoInterfazDef.EnuPorDocumento)
                        If Not HblnEsValido Then
                            HstrMens = "Esta Aplicación Contable solo acepta Interfaz por Documento!"
                            SNotifiqueDatInv()
                        End If
                    End If
                Else
                    HstrMens = "El Tipo de Interfaz seleccionado no es válido!"
                    SNotifiqueDatInv()
                End If
            End If
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If Not e.BlnVaciandoObjeto AndAlso HblnEsValido Then
            MobjPadre.ObjIdAppContableByt.SValide()
        End If
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return ClsOrionCop.FstrNombreDatoConstanteOri(EnuGrupoConstantesOriDef.EnuTipoInterfaz, HobjValorPro)
        End If
    End Function
End Class

Friend Class ClsTipoTerceroCajaByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "TipoTerceroCaja"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Tipo Tercero Caja y Bancos"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsCentroUtilOriCop = ObjPadre
        HblnEsRequerido = lobjPadre.ObjTipoInterfazByt.ObjValorPro <> EnuTipoInterfazDef.None
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew,
                EnuTipoTerceroCajaDef.EnuSinTercero, EnuTipoTerceroCajaDef.EnuCliente,
                BlnEsRequerido)
        If Not HblnEsValido Then
            HstrMens = "El Tipo de Tercero ingresado no es válido!"
            SNotifiqueDatInv()
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(ObjValorPro) Then
            Return ClsOrionCop.FstrNombreDatoConstanteOri(
                    EnuGrupoConstantesOriDef.EnuTipoTercero, HobjValorPro)
        Else
            Return ""
        End If
    End Function
End Class

Friend Class ClsTotalAreaCopropDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "TotalAreaCopropiedad"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "TotalBaseCP"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsCentroUtilOriCop = ObjPadre
        Dim ldecValorMinimo As Decimal = 2D
        If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            ldecValorMinimo = 0
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, ldecValorMinimo, Decimal.MaxValue,
                BlnEsRequerido, HenuTipoValor)
        If HblnEsValido Then
            HobjValorNew = Math.Round(HobjValorNew, 4)
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsTotalAreaPondDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "TotalAreaPonderada"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "TotalBasePoderada"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsCentroUtilOriCop = ObjPadre
        Dim ldecValorMinimo As Decimal = 2D
        If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            ldecValorMinimo = 0
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, ldecValorMinimo,
                Decimal.MaxValue, BlnEsRequerido, HenuTipoValor)
        If HblnEsValido Then
            HobjValorNew = Math.Round(HobjValorNew, 4)
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
#End Region
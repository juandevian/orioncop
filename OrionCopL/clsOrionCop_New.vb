Imports System.Text
Imports System.Runtime.InteropServices
Friend Class ClsOrionCop
#Region "Definiciones"
    Implements IDisposable
    Implements IPanDat
    'Constantes
    Private Const CMSTRCOMA = ", "
    'Variables
    Private MblnDisposed As Boolean = False
    Private ReadOnly MstrArchivoIntegridad As String = GstrTrayDatPrg & "ReporteIntegridad.txt"
    ' Eventos
    Public Event EvnInicio As EventHandler(Of ClsPanEventArgs)
    Public Event EvnAvance As EventHandler(Of ClsPanEventArgs)
    Private MobjArgumentoEventoPan As ClsPanEventArgs = Nothing
    ' Propiedades Autoimplementadas
    Private Property MobjRegistro As Object = Nothing
    Friend Shared Property BlnPreFacturando As Boolean = False
    Friend Shared Property BlnGenFacturas As Boolean = False
    Friend Shared Property BlnFacturando As Boolean = False
    Friend Property BlnCorriendoIntegridad As Boolean = False
    Friend Shared Property BlnProcesoEspecial As Boolean = False
    Friend Shared Property DtmFechaFacturasAReversar As Date = GCDTMFECHANULA
    Private Shared Property DtmFechaFactAuto As Date = GCDTMFECHANULA
    ' Primera CtaCobro Generada
    Friend Shared Property BlnHayCtasCobro As Boolean = False
    ' Integridad
    Private Shared MblnHayerror As Boolean = False
#Region "Ubicación Ventana"
    Const SWP_NOSIZE As Integer = &H1
    Const SWP_NOMOVE As Integer = &H2
    Const SWP_NOACTIVATE As Integer = &H10
    Const wFlags As Integer = SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
#End Region
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Crea una nueva instancia del objeto OrionCop
    ''' </summary>  
    ''' <param name="aenuTipoInstanciamiento">Este argumento indica si la aplicación se esta
    ''' instalando, actualizando o instanciando normalmente.</param>
    ''' <remarks></remarks>
    Public Sub New(aobjRegistro As Object, ablnInicio As Boolean)
        If aobjRegistro Is Nothing OrElse Not (aobjRegistro.GetType.Name = "String") Then
            Throw New ModuloNoRegistradoPanException("El módulo no ha sido debidamente cargado!")
        ElseIf Not aobjRegistro = GCOBJREGISTRO Then
            Throw New ModuloNoRegistradoPanException("El módulo no ha sido debidamente cargado!")
        End If
        MobjRegistro = aobjRegistro
        If ablnInicio Then
            GobjPanorama = New ClsPanorama(ObjRegistro)
            GobjAdministrador = New ClsAdministrador(ObjRegistro)
            GobjPanDat = New ClsPanoramaDat(Me)
        End If
    End Sub
#End Region
#Region "Manejo Tablas Constantes"
    ''' <summary>
    ''' Devuelve una Array de DatRows que contienen los datos de las constates manejadas por el programa generalmente
    ''' a través de enums, según el grupo pasado en el argumento "aenuGrupoTbl"
    ''' </summary>
    ''' <param name="aenuGrupoConstantes">un elemento del enum "enuGrupoConstantesPanDef" que indica el grupo de 
    ''' constantes requerido</param>
    ''' <returns>Array de DataRows</returns>
    ''' <remarks></remarks>
    Friend Shared Function FdrwConstantesOri(aenuGrupoConstantes As EnuGrupoConstantesOriDef) As DataRow()
        Dim lcolTablas As New Collection
        Dim lcolCamposSelect As New Collection
        Dim lcolIndices As New Collection
        lcolTablas.Add("OriTblConstantes")
        lcolCamposSelect.Add({"*"})
        lcolIndices.Add({{"IDGrupo", "ASC"}, {"IDConstante", "ASC"}})
        Dim ldstConstantes As New DataSet
        GobjPanDat.SdsDataSet(ldstConstantes, lcolTablas, lcolCamposSelect, lcolIndices, Nothing)
        Dim ldtbConstantes As DataTable = ldstConstantes.Tables("OriTblConstantes")
        Dim lstrCondicion As String = " IDGrupo = " & CStr(aenuGrupoConstantes)
        Dim ldrwCons As DataRow() = ldtbConstantes.Select(lstrCondicion)
        Return ldrwCons
    End Function
    ''' <summary>
    ''' Devuelve el dato contenido en el campo "Dato" del registro correspondiente al valor pasado en el
    ''' argumento "abytIdDato" del grupo de constantes pasado en el argumento "aenuGrupoTbl"
    ''' </summary>
    ''' <param name="aenuGrupoConstantes">Enu que identifica el grupo de constantes donde se buscara el dato</param>
    ''' <param name="abytIdDato">Id que identifica el dato entre los registros del grupo</param>
    ''' <returns>Nombre de la constante</returns>
    ''' <remarks></remarks>
    Public Shared Function FstrNombreDatoConstanteOri(aenuGrupoConstantes As EnuGrupoConstantesOriDef,
            abytIdDato As Byte)
        Dim ldrwConstates As DataRow() = FdrwConstantesOri(aenuGrupoConstantes)
        For Each ldrwDataRow As DataRow In ldrwConstates
            If ldrwDataRow("IdConstante") = abytIdDato Then
                Return ClsPanorama.FobjValorCampo(ldrwDataRow("Dato"), EnuTipoValor.enuString)
            End If
        Next
        Return ""
    End Function
    ''' <summary>
    ''' Devuelve el dato contenido en el campo "Dato" del registro correspondiente al valor pasado en el
    ''' argumento "abytIdDato" del grupo de constantes pasado en el argumento "aenuGrupoTbl"
    ''' </summary>
    ''' <param name="aenuGrupoConstantes">Enu que identifica el grupo de constantes donde se buscara el dato</param>
    ''' <param name="abytIdDato">Id que identifica el dato entre los registros del grupo</param>
    ''' <returns>Nombre de la constante</returns>
    ''' <remarks></remarks>
#End Region
#Region "Propiedades"
#Region "Propiedades Varias"
    Friend ReadOnly Property ObjArgumentoEventoPan As ClsPanEventArgs
        Get
            If IsNothing(MobjArgumentoEventoPan) Then
                MobjArgumentoEventoPan = New ClsPanEventArgs With {
                    .BlnCancele = False,
                    .BlnVaciandoObjeto = False,
                    .DblCantProcesada = 0.0,
                    .DblCantAProcesar = 0.0
                }
            End If
            Return MobjArgumentoEventoPan
        End Get
    End Property
#End Region
#Region "Clientes"
    ''' <summary>
    ''' Devuelve una nueva instancia de clsCliente.
    ''' </summary>
    ''' <param name="aenuModoInstanciaObj">Define si el objeto Cliente es navegable o unico.</param>
    ''' <remarks>Si la aplicacion no este debidamente registrada se produce una excepcion del tipo
    ''' 'ModuloNoRegistradoPanException'</remarks>
    Friend Shared Function FobjCliente(aenuModoInstanciaObj As EnuModoInstanciaObjDef) As ClsCliente
        Dim lobjCliente As ClsCliente = Nothing
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuNavegable OrElse
                aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuUnico Then
            lobjCliente = New ClsCliente(aenuModoInstanciaObj)
        End If
        Return lobjCliente
    End Function
#End Region
#Region "Predios"
    ''' <summary>
    ''' Devuelve una nueva instancia de clsPredio.
    ''' </summary>
    ''' <param name="aenuModoInstanciaObj">Define si el objeto Predio es navegable o unico.</param>
    ''' <remarks>Si la aplicacion no este debidamente registrada se produce una excepcion del tipo
    ''' 'ModuloNoRegistradoPanException'</remarks>
    Friend Shared Function FobjNuevoPredio(aenuModoInstanciaObj As EnuModoInstanciaObjDef) As ClsPredio
        Dim lobjPredio As ClsPredio = Nothing
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuNavegable OrElse
                aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuUnico Then
            lobjPredio = New ClsPredio(aenuModoInstanciaObj)
        End If
        Return lobjPredio
    End Function
    Friend Shared Function FblnExistePredio(astrIdPredio As String) As Boolean
        Dim lstrFiltro = StrFiltroUbicacion & " AND " & ClsIdPredioStr.SstrNombreCampoBd &
                " = '" & astrIdPredio & "'"
        Dim ldtbPredio = ClsPanorama.FdtbDataTable(ClsPredio.SstrNombreTabla,
                {ClsIdPredio_AntStr.SstrNombreCampoBd}, {{"", ""}}, lstrFiltro)
        Return ldtbPredio.Rows.Count > 0
    End Function
#End Region
#Region "Ubicacion"
    Private Shared ReadOnly Property ObjMiCarpeta As ClsCarpeta
        Get
            Return GobjPanorama.ObjCarpetaActual
        End Get
    End Property
    Friend Shared ReadOnly Property StrFiltroUbicacion As String
        Get
            Return StrCampoCarpeta & " = " & GshrIdCarpeta.ToString & " AND " &
                    StrCampoCentroUtil & " = " &
                    GshrIdCentroUtil.ToString
        End Get
    End Property
    Friend Shared ReadOnly Property StrFiltroUbicacion_Pri As String
        Get
            Dim lstrFiltro = "P." & StrCampoCarpeta & " = " &
                    GshrIdCarpeta.ToString() & " AND P." & StrCampoCentroUtil &
                    " = " & GshrIdCentroUtil.ToString()
            Return lstrFiltro
        End Get
    End Property
    ''' <summary>
    ''' Devuelve la Id. de la Carpeta Actual seguido de su Nombre, separados por un guión.
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Friend Shared ReadOnly Property StrNombreCarpetaActual As String
        Get
            Dim lstrNombreCarActual = String.Empty
            With GobjPanorama.ObjCarpetaActual
                lstrNombreCarActual = .ObjIdCarpetaShr.ToString & " - " & .ObjNombreStr.ObjValorPro
            End With
            Return lstrNombreCarActual
        End Get
    End Property
    ''' <summary>
    ''' Devuelve la Id. de la Copropiedad Actual seguido de su Nombre, separados por un guión.
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Friend Shared ReadOnly Property StrNombreCentroUtilActual As String
        Get
            Dim lstrNombreCenUtilActual = String.Empty
            With GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual
                lstrNombreCenUtilActual = .ObjIdCentroUtilShr.ToString & " - " &
    .ObjNombreCentroUtilStr.ObjValorPro
            End With
            Return lstrNombreCenUtilActual
        End Get
    End Property
#End Region
#Region "Implementa IPanDat"
    Friend ReadOnly Property ObjRegistro As Object Implements IPanDat.ObjRegistro
        Get
            Return MobjRegistro
        End Get
    End Property
    Friend ReadOnly Property ShrIdApp As Short Implements IPanDat.ShrIdApp
        Get
            Return EnuListaAplicaciones.EnuOrionCop
        End Get
    End Property
    Friend ReadOnly Property StrNombreArchivos As String Implements IPanDat.StrNombreArchivos
        Get
            Return "OrionCop_Net"
        End Get
    End Property
    Friend ReadOnly Property EntVersionBDEnProg As Integer Implements IPanDat.EntVersionBDEnProg
        Get
            Return 261
        End Get
    End Property
#End Region
#End Region
#Region "Facturación"
#Region "Generar Prefacturas"
#Region "Prefacturar"
    Friend Shared Function FblnPuedePrefacturar(ByRef astrMens As String) As Boolean
        BlnPreFacturando = True
        Dim lblnPuede = FblnPuedeCrear(EnuTipoDocOri.EnuFactura, False, astrMens)
        If lblnPuede Then
            lblnPuede = Not FblnDebeCausarInt()
            If Not lblnPuede Then
                astrMens = "Antes de generar Pre-Facturas debe causar Intereses de Mora!"
            End If
        End If
        If lblnPuede Then
            lblnPuede = Not FblnHayPrefacturas()
            If Not lblnPuede Then
                astrMens = "Ya hay Pre-Facturas Generadas. No es posible generarlas de nuevo!"
            End If
        End If
        If lblnPuede Then
            lblnPuede = FblnHayItemsPorFacturar()
            If Not lblnPuede Then
                astrMens = "No hay Facturas programadas para procesar!"
            End If
        End If
        If lblnPuede Then
            lblnPuede = GobjParametros.FblnServicioIdOk(astrMens)
        End If
        If lblnPuede Then
            lblnPuede = GobjParametros.FblnParaCuotaAdminOk(astrMens)
        End If
        If lblnPuede Then
            lblnPuede = Not GobjParametros.ObjAnoActual.FblnDebeCalcularCuotas
            If Not lblnPuede Then
                astrMens = "Antes de generar Pre-Facturas debe calcular las Cuotas de Administración!"
            End If
        End If
        If lblnPuede Then
            lblnPuede = Not GobjParametros.ObjAnoActual.FblnDebeAjustarCuotasAdmin
            If Not lblnPuede Then
                astrMens = "Antes de generar Pre-Facturas debe ajustar las Cuotas de Administración!"
            End If
        End If
        If lblnPuede Then
            lblnPuede = Not FblnServicioSinGenerarItemsPrograma()
            If Not lblnPuede Then
                astrMens = "Hay Programación de Facturas por llevarse a cabo. No es posible Pre-Facturar!"
            End If
        End If
        If lblnPuede Then
            lblnPuede = FblnEmailOk(astrMens)
            If Not lblnPuede Then
                astrMens = "Hay Clientes con el Email desconfigurado. No es posible Pre-Facturar!"
            End If
        End If
        BlnPreFacturando = False
        Return lblnPuede
    End Function
    ''' <summary>
    ''' Indica si en el momento existe algún servicio que genera programa de facturación, 
    ''' que aún no haya generado los items del programa.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function FblnServicioSinGenerarItemsPrograma() As Boolean
        Dim lblnHay = False
        For Each lobjServicio As ClsServicio In GobjParametros.ObjAnoActual.ColServiciosAno
            If lobjServicio.ObjGeneraProgramBln.ObjValorPro Then
                lblnHay = Not lobjServicio.ObjEstaGenaradaProgramBln.ObjValorPro
                If lblnHay Then Exit For
            End If
        Next
        If Not lblnHay Then
            For Each lobjServicio As ClsServicio In GobjParametros.ColServiciosPer
                If lobjServicio.ObjGeneraProgramBln.ObjValorPro Then
                    lblnHay = Not lobjServicio.ObjEstaGenaradaProgramBln.ObjValorPro
                    If lblnHay Then Exit For
                End If
            Next
        End If
        Return lblnHay
    End Function
    Friend Sub SGenerePrefacturas()
        Dim lblnNoHayError = False
        GobjPanDat.SControleProcesoObj(True)
        Try
            BlnPreFacturando = True
            MobjArgumentoEventoPan = Nothing
            GobjPanDat.SInicialiceTransaccion()
            Dim larlFechasAFactHoy = GobjParametros.FarlFechasAFActurarHoy
            For Each ldtmFechaAFac As Date In larlFechasAFactHoy
                SPrefacture(ldtmFechaAFac)
            Next
            GobjPanorama.SRegistreAccionLogApp("clsOrionCop", "Genera Pre-facturas")
            lblnNoHayError = True
        Catch ex As ErrorInesperadoPanDatException
            Throw
        Catch ex As ErrorInesperadoPanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As PanLException
            Throw
        Catch ex As ArgumentException
            Throw
        Catch ex As Exception
            Throw
        Finally
            BlnPreFacturando = False
            If Not lblnNoHayError OrElse ObjArgumentoEventoPan.BlnCancele Then
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
                If ObjArgumentoEventoPan.BlnCancele Then
                    ObjArgumentoEventoPan.EnuProceso = EnuProcesoDef.None
                    ObjArgumentoEventoPan.DblCantAProcesar = 0
                    ObjArgumentoEventoPan.DblCantProcesada = 0
                    RaiseEvent EvnInicio(Me, ObjArgumentoEventoPan)
                End If
            Else
                GobjPanDat.SConfirmeTransaccion()
                GobjPanDat.SControleProcesoObj(False)
            End If
        End Try
    End Sub
    ''' <summary>
    ''' Indica si estan dadas las condiciones para generara las Pre-Facturas
    ''' </summary>
    ''' <param name="astrMens">En este argumento se devuelve el mensaje que indica el motivo
    ''' que no permite prefacturar</param>
    ''' <returns></returns>
    Private Sub SPrefacture(adtmFechaFac As Date)
        Dim ldtbItemsProgAPreFact = FdtbItemsAPreFactPredios()
        Dim lstrIdPredAgrAFac = FstrPredAgrAFac(ldtbItemsProgAPreFact)
        ' Prefacturar Predios servicios solo a Propietarios
        SPrefPrediosAgrAProp(adtmFechaFac, ldtbItemsProgAPreFact, lstrIdPredAgrAFac)
        ' Prefacturar Predios servicios no solo a Propietarios
        SPrefAPredios(adtmFechaFac, ldtbItemsProgAPreFact, lstrIdPredAgrAFac)
        ' Prefacturar a clientes
        ldtbItemsProgAPreFact = FdtbItemsAPreFactClientes()
        Dim ldblClientesAFac = FdblClientesAFac(ldtbItemsProgAPreFact)
        SPrefAClientes(adtmFechaFac, ldtbItemsProgAPreFact, ldblClientesAFac)
    End Sub
    Private Sub SPrefPrediosAgrAProp(adtmFecha As Date, adtbItemsProgAFac As DataTable,
            astrIdPredsAgrAFac As List(Of String))
        Dim lobjPredAgr As New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
        Dim lobjValorLlave As Object(), ldrwItemsAFac As DataRow()
        Dim lstrKeysSerAFac = FstrServiciosAFact(adtmFecha, True), lstrfiltro As String
        If lstrKeysSerAFac.Count > 0 Then
            Dim i = 0
            ObjArgumentoEventoPan.BlnCancele = False
            ObjArgumentoEventoPan.DblCantAProcesar = astrIdPredsAgrAFac.Count
            ObjArgumentoEventoPan.EnuProceso = EnuProcesoDef.EnuPreFacProPA
            ObjArgumentoEventoPan.DblCantProcesada = 0
            RaiseEvent EvnInicio(Me, ObjArgumentoEventoPan)
            RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
            For Each lstrIdpreAFac As String In astrIdPredsAgrAFac
                i += 1
                lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrIdpreAFac}
                lobjPredAgr.SAbra(lobjValorLlave)
                lstrfiltro = FstrFiltroPredYServicios(lstrIdpreAFac, lstrKeysSerAFac)
                ldrwItemsAFac = adtbItemsProgAFac.Select(lstrfiltro)
                If ldrwItemsAFac.Length > 0 Then
                    If GobjParametros.ObjConsolidaItemsFacBln.ObjValorPro AndAlso
                            Not lobjPredAgr.ObjNoConsolidarItemsFacBln.ObjValorPro Then
                        SPreFactureAPredioAgrConsol(lobjPredAgr, ldrwItemsAFac,
                                adtmFecha, True)
                    Else
                        SPreFactureAPredio(lobjPredAgr, ldrwItemsAFac, adtmFecha, True)
                    End If
                    SActualiceItemsProgFac(ldrwItemsAFac, True)
                End If
                ObjArgumentoEventoPan.BlnCancele = False
                ObjArgumentoEventoPan.DblCantProcesada = i
                RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
                If ObjArgumentoEventoPan.BlnCancele Then
                    Exit For
                End If
            Next
        End If
    End Sub
    Private Sub SPrefAPredios(adtmfechaFac As Date, adtbItemsProgAFac As DataTable,
            astrIdPredsAgrAFac As List(Of String))
        Dim lobjPredAgr As New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
        Dim lobjValorLlave As Object(), ldrwItemsAFac As DataRow()
        Dim lstrKeysSerAFac = FstrServiciosAFact(adtmfechaFac, False)
        Dim lstrfiltro As String, i = 0
        Dim lstrOrden = ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " DESC"
        If lstrKeysSerAFac.Count > 0 Then
            ObjArgumentoEventoPan.EnuProceso = EnuProcesoDef.enuPreFacPre
            ObjArgumentoEventoPan.DblCantAProcesar = astrIdPredsAgrAFac.Count
            ObjArgumentoEventoPan.DblCantProcesada = i
            RaiseEvent EvnInicio(Me, ObjArgumentoEventoPan)
            RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
            For Each lstrIdpreAFac As String In astrIdPredsAgrAFac
                i += 1
                lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrIdpreAFac}
                lobjPredAgr.SAbra(lobjValorLlave)
                If lobjPredAgr.ObjFacturarPorServicio_PreBln.ObjValorPro Then
                    For Each lstrKeySer As String In lstrKeysSerAFac
                        lstrfiltro = FstrFiltroPredYServicio(lstrIdpreAFac, lstrKeySer)
                        ldrwItemsAFac = adtbItemsProgAFac.Select(lstrfiltro, lstrOrden)
                        If ldrwItemsAFac.Length > 0 Then
                            If GobjParametros.ObjConsolidaItemsFacBln.ObjValorPro AndAlso
                                    Not lobjPredAgr.ObjNoConsolidarItemsFacBln.
                                    ObjValorPro Then
                                SPreFactureAPredioAgrConsol(lobjPredAgr, ldrwItemsAFac,
                                        adtmfechaFac, False)
                            Else
                                SPreFactureAPredio(lobjPredAgr, ldrwItemsAFac,
                                        adtmfechaFac, False)
                            End If
                            SActualiceItemsProgFac(ldrwItemsAFac, True)
                        End If
                    Next
                Else
                    lstrfiltro = FstrFiltroPredYServicios(lstrIdpreAFac, lstrKeysSerAFac)
                    ldrwItemsAFac = adtbItemsProgAFac.Select(lstrfiltro, lstrOrden)
                    If ldrwItemsAFac.Length > 0 Then
                        If GobjParametros.ObjConsolidaItemsFacBln.ObjValorPro AndAlso
                                Not lobjPredAgr.ObjNoConsolidarItemsFacBln.ObjValorPro Then
                            SPreFactureAPredioAgrConsol(lobjPredAgr, ldrwItemsAFac,
                                    adtmfechaFac, False)
                        Else
                            SPreFactureAPredio(lobjPredAgr, ldrwItemsAFac, adtmfechaFac,
                                    False)
                        End If
                        SActualiceItemsProgFac(ldrwItemsAFac, True)
                    End If
                End If
                ObjArgumentoEventoPan.DblCantProcesada = i
                ObjArgumentoEventoPan.BlnCancele = False
                RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
                If ObjArgumentoEventoPan.BlnCancele Then
                    Exit For
                End If
            Next
        End If
    End Sub
    Private Shared Sub SPreFactureAPredio(aobjPredioAgr As ClsPredio,
            adrwItemsAFact As DataRow(), adtmFecha As Date, ablnAPropietario As Boolean)
        Dim lobjCliente As ClsCliente = Nothing
        Dim lcolFacturas As New Collection, lobjFactura As ClsFactura
        Dim lstrIdPredAgr = aobjPredioAgr.ObjIdPredioAgrupadorStr.ToString()
        If aobjPredioAgr.ObjIdTipoDestinatarioFacturaByt.ObjValorPro =
                EnuDestinatarioFacturaDef.EnuPropietario OrElse ablnAPropietario Then
            For Each lobjProp As ClsPropietario In aobjPredioAgr.ColPropietarios
                lobjFactura = FobjNuevaFact(lobjProp.ObjCliente, lstrIdPredAgr,
                        adtmFecha)
                lcolFacturas.Add(lobjFactura, lobjProp.ObjIdCliente_PropDbl.ToString())
            Next
        Else
            lobjCliente = aobjPredioAgr.ObjArrendatario
            lobjFactura = FobjNuevaFact(lobjCliente, lstrIdPredAgr, adtmFecha)
            lcolFacturas.Add(lobjFactura, lobjCliente.ObjIdClienteDbl.ToString())
        End If
        Dim ldecVlrItem As Decimal, lcolValorItemProps As Collection
        Dim lstrIdPredio As String
        Dim lstrKeySer As String, lstrIdAno As String, lstrIdServicio As String
        For Each ldrwItemAFac As DataRow In adrwItemsAFact
            lstrIdAno = ClsPanorama.FobjValorCampo(ldrwItemAFac(
                    ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            lstrIdServicio = ClsPanorama.FobjValorCampo(ldrwItemAFac(
                    ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            lstrKeySer = lstrIdAno & "," & lstrIdServicio
            lstrIdPredio = ClsPanorama.FobjValorCampo(ldrwItemAFac(
                    ClsIdPredio_ItemProgramaFactStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            ldecVlrItem = ClsPanorama.FobjValorCampo(ldrwItemAFac(
                    ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd),
                    EnuTipoValor.enuDecimal)
            lcolValorItemProps = FcolVlrItemPropietarios(aobjPredioAgr, ldecVlrItem,
                    ablnAPropietario)
            If aobjPredioAgr.ObjIdTipoDestinatarioFacturaByt.ObjValorPro =
                    EnuDestinatarioFacturaDef.EnuPropietario OrElse ablnAPropietario Then
                SInserteItemsFact(aobjPredioAgr.ColPropietarios, lcolFacturas,
                        lcolValorItemProps, lstrIdPredio, lstrKeySer)
            Else
                SInserteItemsFact(lobjCliente, lcolFacturas, lcolValorItemProps,
                        lstrIdPredio, lstrKeySer)
            End If
        Next
        For Each lobjFact As ClsFactura In lcolFacturas
            lobjFact.SActualice(True)
        Next
    End Sub
    Private Shared Sub SPreFactureAPredioAgrConsol(aobjPredioAgr As ClsPredio,
            adrwItemsAFact As DataRow(), adtmFecha As Date, ablnAPropietario As Boolean)
        Dim lobjCliente As ClsCliente = Nothing
        Dim lcolFacturas As New Collection, lobjFactura As ClsFactura
        Dim lstrIdPredAgr = aobjPredioAgr.ObjIdPredioAgrupadorStr.ToString()
        If aobjPredioAgr.ObjIdTipoDestinatarioFacturaByt.ObjValorPro =
                EnuDestinatarioFacturaDef.EnuPropietario OrElse ablnAPropietario Then
            For Each lobjProp As ClsPropietario In aobjPredioAgr.ColPropietarios
                lobjFactura = FobjNuevaFact(lobjProp.ObjCliente, lstrIdPredAgr,
                        adtmFecha)
                lcolFacturas.Add(lobjFactura, lobjProp.ObjIdCliente_PropDbl.ToString())
            Next
        Else
            lobjCliente = aobjPredioAgr.ObjArrendatario
            lobjFactura = FobjNuevaFact(lobjCliente, lstrIdPredAgr, adtmFecha)
            lcolFacturas.Add(lobjFactura, lobjCliente.ObjIdClienteDbl.ToString())
        End If
        Dim ldecVlrItem As Decimal, lcolValorItemProps As Collection
        Dim lstrKeySer = String.Empty, lstrIdAno As String, lstrIdServicio As String
        For Each ldrwItemAFac As DataRow In adrwItemsAFact
            lstrIdAno = ClsPanorama.FobjValorCampo(ldrwItemAFac(
                    ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            lstrIdServicio = ClsPanorama.FobjValorCampo(ldrwItemAFac(
                    ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            If lstrKeySer <> lstrIdAno & "," & lstrIdServicio Then
                If ldecVlrItem > 0 Then
                    lcolValorItemProps = FcolVlrItemPropietarios(aobjPredioAgr,
                            ldecVlrItem, ablnAPropietario)
                    If aobjPredioAgr.ObjIdTipoDestinatarioFacturaByt.ObjValorPro =
                            EnuDestinatarioFacturaDef.EnuPropietario OrElse
                            ablnAPropietario Then
                        SInserteItemsFact(aobjPredioAgr.ColPropietarios, lcolFacturas,
                                lcolValorItemProps, lstrIdPredAgr, lstrKeySer)
                    Else
                        SInserteItemsFact(lobjCliente, lcolFacturas, lcolValorItemProps,
                                lstrIdPredAgr, lstrKeySer)
                    End If
                    ldecVlrItem = 0
                End If
                lstrKeySer = lstrIdAno & "," & lstrIdServicio
                ldecVlrItem += ClsPanorama.FobjValorCampo(ldrwItemAFac(
                        ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd),
                        EnuTipoValor.enuDecimal)
            Else
                ldecVlrItem += ClsPanorama.FobjValorCampo(ldrwItemAFac(
                        ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd),
                        EnuTipoValor.enuDecimal)
            End If
        Next
        If ldecVlrItem > 0 Then
            lcolValorItemProps = FcolVlrItemPropietarios(aobjPredioAgr,
                            ldecVlrItem, ablnAPropietario)
            If aobjPredioAgr.ObjIdTipoDestinatarioFacturaByt.ObjValorPro =
                            EnuDestinatarioFacturaDef.EnuPropietario OrElse
                            ablnAPropietario Then
                SInserteItemsFact(aobjPredioAgr.ColPropietarios, lcolFacturas,
                                lcolValorItemProps, lstrIdPredAgr, lstrKeySer)
            Else
                SInserteItemsFact(lobjCliente, lcolFacturas, lcolValorItemProps,
                                lstrIdPredAgr, lstrKeySer)
            End If
        End If
        For Each lobjFact As ClsFactura In lcolFacturas
            lobjFact.SActualice(True)
        Next
    End Sub
    Private Sub SPrefAClientes(adtmFechaFac As Date,
            adtbItemsProgAFac As DataTable, adblClientesAFac As List(Of Double))
        Dim lobjCliente As New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
        Dim lobjValorLlave As Object(), ldrwItemsAFac As DataRow()
        Dim lstrKeysSerAFac = FstrServiciosAFact(adtmFechaFac, False)
        Dim lstrfiltro As String, i = 0
        Dim lstrOrden = ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd & " ASC"
        If lstrKeysSerAFac.Count > 0 Then
            ObjArgumentoEventoPan.EnuProceso = EnuProcesoDef.enuPreFacCli
            ObjArgumentoEventoPan.DblCantAProcesar = adblClientesAFac.Count
            ObjArgumentoEventoPan.DblCantProcesada = i
            RaiseEvent EvnInicio(Me, ObjArgumentoEventoPan)
            RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
            For Each ldblIdCliente As Double In adblClientesAFac
                i += 1
                lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, ldblIdCliente}
                lobjCliente.SAbra(lobjValorLlave)
                If lobjCliente.ObjFactPorServicio_CliBln.ObjValorPro Then
                    For Each lstrKeySer As String In lstrKeysSerAFac
                        If lstrKeySer.Split(",")(0) = 0 Then
                            lstrfiltro = FstrFiltroCliYServicio(ldblIdCliente, lstrKeySer)
                            ldrwItemsAFac = adtbItemsProgAFac.Select(lstrfiltro, lstrOrden)
                            If ldrwItemsAFac.Length > 0 Then
                                SPreFactureACliente(lobjCliente, ldrwItemsAFac, adtmFechaFac)
                                SActualiceItemsProgFac(ldrwItemsAFac, False)
                            End If
                        End If
                    Next
                Else
                    lstrfiltro = FstrFiltroClieYServicios(ldblIdCliente, lstrKeysSerAFac)
                    If Not String.IsNullOrEmpty(lstrfiltro) Then
                        ldrwItemsAFac = adtbItemsProgAFac.Select(lstrfiltro, lstrOrden)
                        If ldrwItemsAFac.Length > 0 Then
                            SPreFactureACliente(lobjCliente, ldrwItemsAFac, adtmFechaFac)
                            SActualiceItemsProgFac(ldrwItemsAFac, False)
                        End If
                    End If
                End If
                ObjArgumentoEventoPan.BlnCancele = False
                ObjArgumentoEventoPan.DblCantProcesada = i
                RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
                If ObjArgumentoEventoPan.BlnCancele Then
                    Exit For
                End If
            Next
        End If
    End Sub
    Private Shared Sub SPreFactureACliente(aobjCliente As ClsCliente,
            adrwItemsAFact As DataRow(), adtmFecha As Date)
        Dim lobjFactura = FobjNuevaFact(aobjCliente, "", adtmFecha)
        Dim ldblIdCliente As Double = aobjCliente.ObjIdClienteDbl.ObjValorPro
        Dim ldecVlrItem As Decimal
        Dim lstrKeySer As String, lstrIdAno As String, lstrIdServicio As String
        For Each ldrwItemAFac As DataRow In adrwItemsAFact
            lstrIdAno = ClsPanorama.FobjValorCampo(ldrwItemAFac(
                    ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            lstrIdServicio = ClsPanorama.FobjValorCampo(ldrwItemAFac(
                    ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            ldblIdCliente = ClsPanorama.FobjValorCampo(ldrwItemAFac(
                    ClsIdCliente_ItemProgramaFactDbl.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            ldecVlrItem = ClsPanorama.FobjValorCampo(ldrwItemAFac(
                    ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd),
                    EnuTipoValor.enuDecimal)
            lstrKeySer = lstrIdAno & "," & lstrIdServicio
            SAddItemFact(lobjFactura, "", lstrKeySer, ldecVlrItem)
        Next
        lobjFactura.SActualice(True)
    End Sub
    Private Shared Sub SInserteItemsFact(acolProps As Collection,
            acolFacturas As Collection, acolVlrItemProps As Collection,
            astrIdPredio As String, astrKeySer As String)
        Dim ldecVlrItemProp As Decimal, lobjFact As ClsFactura, lstrIdCliente As String
        For Each lobjProp As ClsPropietario In acolProps
            lstrIdCliente = lobjProp.ObjIdCliente_PropDbl.ToString()
            ldecVlrItemProp =
                    CType(acolVlrItemProps(lstrIdCliente).ToString().Split(",")(0), Decimal)
            If ldecVlrItemProp > 0 Then
                lobjFact = acolFacturas(lobjProp.ObjIdCliente_PropDbl.ToString())
                SAddItemFact(lobjFact, astrIdPredio, astrKeySer, ldecVlrItemProp)
            End If
        Next
    End Sub
    Private Shared Sub SInserteItemsFact(aobjCliente As ClsCliente,
            acolFacturas As Collection, acolVlrItemProps As Collection,
            astrIdPredio As String, astrKeySer As String)
        Dim ldecVlrItemProp As Decimal, lobjFact As ClsFactura, lstrIdCliente As String
        lstrIdCliente = aobjCliente.ObjIdClienteDbl.ToString()
        ldecVlrItemProp =
                CType(acolVlrItemProps(lstrIdCliente).ToString().Split(",")(0), Decimal)
        If ldecVlrItemProp > 0 Then
            lobjFact = acolFacturas(aobjCliente.ObjIdClienteDbl.ToString())
            SAddItemFact(lobjFact, astrIdPredio, astrKeySer, ldecVlrItemProp)
        End If
    End Sub
    Private Shared Sub SAddItemFact(aobjFactura As ClsFactura, astrIdPredio As String,
            astrKeyServ As String, adecValor As Decimal)
        Dim lshrIdAno = CShort(astrKeyServ.Split(",")(0))
        Dim lshrIdServ = CShort(astrKeyServ.Split(",")(1))
        Dim lobjServicio As ClsServicio = GobjParametros.FobjServicio(astrKeyServ)
        Dim lstrDetalle = lobjServicio.ObjNombreServicioStr.ObjValorPro & " " & astrIdPredio
        If lobjServicio.BlnEsCuotaAdministracion Then
            If lobjServicio.ObjEsAjusteBln.ObjValorPro Then
                lstrDetalle &= " " & GobjParametros.ObjAnoActual.ObjIdAnoShr.ToString
            Else
                lstrDetalle &= " " & GobjParametros.ObjAnoActual.StrNombrePeriodoActual
            End If
        End If
        Dim lobjItemFact As ClsItemFactura = aobjFactura.FobjNuevoItemFactura
        With lobjItemFact
            .ObjFechaGraciaIFDtm.ObjValorPro = lobjServicio.DtmFechaGraciaPeriActual
            .ObjFechaVencimientoIFDtm.ObjValorPro = lobjServicio.DtmFechaVencePeriActual
            .ObjEsPrefactura_ItemFactBln.ObjValorPro = aobjFactura.ObjEsPreFacturaBln.ObjValorPro
            .ObjDetalle_ItemFactStr.ObjValorPro = lstrDetalle
            .ObjIdAno_ServicioItemFactShr.ObjValorPro = lshrIdAno
            .ObjIdServicio_ItemFactShr.ObjValorPro = lshrIdServ
            .ObjIdPredio_ItemFactStr.ObjValorPro = astrIdPredio
            .ObjValor_ItemFactDec.ObjValorPro = adecValor
        End With
        aobjFactura.SAdicioneNuevoItem(lobjItemFact)
    End Sub
    Private Shared Function FdtbItemsAPreFactPredios() As DataTable
        Dim lshrIdAnoActual As Short =
                GobjParametros.ObjAnoActual.ObjIdAnoShr.ObjValorPro
        Dim lshrIdMesAct As Short =
                GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo.Month
        Dim lstrSql = "SELECT DISTINCT UPPER(" &
                ClsIdPredioAgrupadorStr.SstrNombreCampoBd & ") AS IdPreAgr, I." &
                ClsIdPredio_ItemProgramaFactStr.SstrNombreCampoBd & ", " &
                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & ", " &
                ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd & ", " &
                ClsNoConsolidarItemsFacBln.SstrNombreCampoBd & ", " &
                ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd & " FROM " &
                ClsPredio.SstrNombreTabla & " AS P INNER JOIN (SELECT " &
                StrCampoCarpeta & ", " &
                StrCampoCentroUtil & ", " &
                ClsIdPredio_ItemProgramaFactStr.SstrNombreCampoBd & ", " &
                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & ", " &
                ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd & ", " &
                ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd & " FROM " &
                ClsItemProgramaFact.SstrNombreTabla & " WHERE " & StrFiltroUbicacion &
                " AND " & ClsSaldo_ItemProgramaFactDec.SstrNombreCampoBd & " > 0 AND " &
                ClsIdPredio_ItemProgramaFactStr.SstrNombreCampoBd & " <> '' AND 
                ((SUBSTRING(" & ClsPeriodoIni_ItemProgStr.SstrNombreCampoBd &
                ",1,4) - 1) * 12) + SUBSTRING(" &
                ClsPeriodoIni_ItemProgStr.SstrNombreCampoBd & ",5,2) + " &
                ClsCantidadPeriodosShr.SstrNombreCampoBd & "  - (" &
                ClsSaldo_ItemProgramaFactDec.SstrNombreCampoBd & "/" &
                ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd & ") - 1 < ((" &
                lshrIdAnoActual & "-1)*12) + " & lshrIdMesAct & ") AS I ON P." &
                StrCampoCarpeta & " = I." &
                StrCampoCarpeta & " AND P." &
                StrCampoCentroUtil & " = I." &
                StrCampoCentroUtil & " AND P." &
                ClsIdPredioStr.SstrNombreCampoBd & " = I." &
                ClsIdPredio_ItemProgramaFactStr.SstrNombreCampoBd & " ORDER BY " &
                ClsIdPredioAgrupadorStr.SstrNombreCampoBd & ", " &
                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " DESC, " &
                ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd & ", " &
                ClsIdPredio_ItemProgramaFactStr.SstrNombreCampoBd
        Dim ldtbRes = ClsPanorama.FdtbDataTable(lstrSql)
        Return ldtbRes
    End Function
    ''' <summary>
    ''' Devuelve una datatable donde se consolidad los valores de los items a facturar
    ''' de los predios en el predio agrupador
    ''' </summary>
    ''' <returns></returns>
    Private Shared Function FdtbItemsAPreFactClientes() As DataTable
        Dim lstrTabla = ClsItemProgramaFact.SstrNombreTabla
        Dim lshrIdAnoActual As Short = GobjParametros.ObjAnoActual.ObjIdAnoShr.ObjValorPro
        Dim lshrIdMesAct As Short =
                GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo.Month
        Dim lstrSql = "SELECT " &
                ClsIdCliente_ItemProgramaFactDbl.SstrNombreCampoBd & ", " &
                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & ", " &
                ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd & ", " &
                ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd & " FROM " &
                ClsItemProgramaFact.SstrNombreTabla & " WHERE " &
                StrFiltroUbicacion & " AND " &
                ClsSaldo_ItemProgramaFactDec.SstrNombreCampoBd & " > 0 AND " &
                ClsIdPredio_ItemProgramaFactStr.SstrNombreCampoBd & " = '' AND ((SUBSTRING(" &
                ClsPeriodoIni_ItemProgStr.SstrNombreCampoBd & ",1,4) - 1) * 12) + SUBSTRING(" &
                ClsPeriodoIni_ItemProgStr.SstrNombreCampoBd & ",5,2) + " &
                ClsCantidadPeriodosShr.SstrNombreCampoBd & "  - (" &
                ClsSaldo_ItemProgramaFactDec.SstrNombreCampoBd & "/" &
                ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd & ") - 1 <  ((" &
                lshrIdAnoActual & "-1)*12) + " & lshrIdMesAct & " ORDER BY " &
                ClsIdCliente_ItemProgramaFactDbl.SstrNombreCampoBd & ", " &
                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " DESC, " &
                ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd
        Dim ldtbRes = ClsPanorama.FdtbDataTable(lstrSql)
        Return ldtbRes
    End Function
    Private Shared Sub SActualiceItemsProgFac(adrwItemsProgFac As DataRow(),
        ablnPredio As Boolean)
        Dim lstrTabla = ClsItemProgramaFact.SstrNombreTabla
        Dim lstrIdPredio = String.Empty, ldecVlr As Decimal, ldblIdCliente = 0.0
        Dim lshrIdAno As Short, lshrIdSer As Short
        Dim lstrExpSql As String, lstrFiltro As String, lstrFiltroDest As String
        For Each ldrwItem As DataRow In adrwItemsProgFac
            If ablnPredio Then
                lstrIdPredio = ClsPanorama.FobjValorCampo(ldrwItem(
                    ClsIdPredio_ItemProgramaFactStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            Else
                ldblIdCliente = ClsPanorama.FobjValorCampo(ldrwItem(
                    ClsIdCliente_ItemProgramaFactDbl.SstrNombreCampoBd),
                    EnuTipoValor.enuDouble)
            End If
            ldecVlr = ClsPanorama.FobjValorCampo(ldrwItem(
                    ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd),
                    EnuTipoValor.enuDecimal)
            lshrIdAno = ClsPanorama.FobjValorCampo(ldrwItem(
                    ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd),
                    EnuTipoValor.enuShort)
            lshrIdSer = ClsPanorama.FobjValorCampo(ldrwItem(
                    ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd),
                    EnuTipoValor.enuShort)
            lstrExpSql = "UPDATE " & lstrTabla & " SET " &
                    ClsSaldo_ItemProgramaFactDec.SstrNombreCampoBd & " = " &
                    ClsSaldo_ItemProgramaFactDec.SstrNombreCampoBd & " - " &
                    ldecVlr
            If ablnPredio Then
                lstrFiltroDest = ClsIdPredio_ItemProgramaFactStr.SstrNombreCampoBd &
                        " = '" & lstrIdPredio & "'"
            Else
                lstrFiltroDest = ClsIdCliente_ItemProgramaFactDbl.SstrNombreCampoBd &
                        " = " & ldblIdCliente
            End If
            lstrFiltro = StrFiltroUbicacion & " AND " & lstrFiltroDest & " AND " &
                    ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " = " &
                    lshrIdAno & " AND " &
                    ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd & " = " &
                    lshrIdSer
            lstrExpSql &= " WHERE " & lstrFiltro
            GobjPanDat.SEjecuteSentenciaSql(lstrExpSql)
        Next
    End Sub
    Private Shared Function FstrServiciosAFact(adtmFechaFac As Date,
            ablnSoloProp As Boolean) As List(Of String)
        Dim lstrKeySer As New List(Of String)
        For Each lobjSer As ClsServicio In GobjParametros.ObjAnoActual.ColServiciosAno
            If lobjSer.DtmFechaFacturacionPeriodoActual = adtmFechaFac Then
                If ablnSoloProp Then
                    If lobjSer.ObjFactAPropYPreAgrBln.ObjValorPro Then
                        lstrKeySer.Add(lobjSer.FstrMyKey)
                        Exit For
                    End If
                Else
                    If Not lobjSer.ObjFactAPropYPreAgrBln.ObjValorPro Then
                        lstrKeySer.Add(lobjSer.FstrMyKey)
                    End If
                End If
            End If
        Next
        For Each lobjSer As ClsServicio In GobjParametros.ColServiciosPer
            If lobjSer.DtmFechaFacturacionPeriodoActual = adtmFechaFac Then
                If ablnSoloProp Then
                    If lobjSer.ObjFactAPropYPreAgrBln.ObjValorPro Then
                        lstrKeySer.Add(lobjSer.FstrMyKey)
                    End If
                Else
                    If Not lobjSer.ObjFactAPropYPreAgrBln.ObjValorPro Then
                        lstrKeySer.Add(lobjSer.FstrMyKey)
                    End If
                End If
            End If
        Next
        Return lstrKeySer
    End Function
    Private Shared Function FstrPredAgrAFac(adtbItemsAFac As DataTable) As List(Of String)
        Dim lstrPredAFac As New List(Of String), lstrIdpreAgr As String
        For Each ldrwITemAFac As DataRow In adtbItemsAFac.Rows
            lstrIdpreAgr = ClsPanorama.FobjValorCampo(ldrwITemAFac("IdPreAgr"),
                    EnuTipoValor.enuString)
            If Not lstrPredAFac.Contains(lstrIdpreAgr) Then
                lstrPredAFac.Add(lstrIdpreAgr)
            End If
        Next
        Return lstrPredAFac
    End Function
    Private Shared Function FdblClientesAFac(adtbItemsAFac As DataTable) As List(Of Double)
        Dim ldblCliAFac As New List(Of Double), ldblIdCliente As Double
        For Each ldrwITemAFac As DataRow In adtbItemsAFac.Rows
            ldblIdCliente = ClsPanorama.FobjValorCampo(ldrwITemAFac(
                    ClsIdCliente_ItemProgramaFactDbl.SstrNombreCampoBd),
                    EnuTipoValor.enuDouble)
            If Not ldblCliAFac.Contains(ldblIdCliente) Then
                ldblCliAFac.Add(ldblIdCliente)
            End If
        Next
        ldblCliAFac.Sort()
        Return ldblCliAFac
    End Function
    Private Shared Function FstrFiltroPredYServicios(astrIdpredAgr As String,
            astrKeysSer As List(Of String)) As String
        Dim lstrFiltro = "IdPreAgr = '" & astrIdpredAgr & "' AND (("
        Dim lstrIdAno As String, lstrIdSer As String, i = 0
        For Each lstrKeySer As String In astrKeysSer
            lstrIdAno = lstrKeySer.Split(",")(0)
            lstrIdSer = lstrKeySer.Split(",")(1)
            lstrFiltro &= ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " = " &
                    lstrIdAno & " AND " &
                    ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd & " = " &
                    lstrIdSer & ")"
            i += 1
            If i < astrKeysSer.Count Then
                lstrFiltro &= " OR ("
            Else
                lstrFiltro &= ")"
            End If
        Next
        Return lstrFiltro
    End Function
    Private Shared Function FstrFiltroPredYServicio(astrIdpredAgr As String,
            astrKeySer As String) As String
        Dim lstrIdAno As String, lstrIdSer As String, lstrFiltro As String
        lstrIdAno = astrKeySer.Split(",")(0)
        lstrIdSer = astrKeySer.Split(",")(1)
        If lstrIdAno = "0" Then
            lstrFiltro = "IdPreAgr = '" & astrIdpredAgr & "' AND " &
                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " = " & lstrIdAno &
                " AND " & ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd &
                " = " & lstrIdSer
        Else
            lstrFiltro = "IdPreAgr = '" & astrIdpredAgr & "' AND " &
                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " = " & lstrIdAno &
                " AND " & ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd &
                " = " & lstrIdSer
        End If
        Return lstrFiltro
    End Function
    Private Shared Function FstrFiltroClieYServicios(adblIdCliente As Double,
            astrKeysSer As List(Of String)) As String
        Dim lstrFiltro = ClsIdCliente_ItemProgramaFactDbl.SstrNombreCampoBd & " = " &
                adblIdCliente & " AND ((", lstrFiltroSer = String.Empty
        Dim lstrIdAno As String, lstrIdSer As String, i = 0
        For Each lstrKeySer As String In astrKeysSer
            lstrIdAno = lstrKeySer.Split(",")(0)
            lstrIdSer = lstrKeySer.Split(",")(1)
            If lstrIdAno = 0 Then
                lstrFiltroSer &= ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " = " &
                        lstrIdAno & " AND " &
                        ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd & " = " &
                        lstrIdSer & ")"
                i += 1
                If i < astrKeysSer.Count Then
                    lstrFiltroSer &= " OR ("
                Else
                    lstrFiltroSer &= ")"
                End If
            Else
                i += 1
            End If
        Next
        If String.IsNullOrEmpty(lstrFiltroSer) Then
            lstrFiltro = String.Empty
        Else
            lstrFiltro &= lstrFiltroSer
        End If
        Return lstrFiltro
    End Function
    Private Shared Function FstrFiltroCliYServicio(adblIdCliente As Double,
            astrKeySer As String) As String
        Dim lstrIdAno As String, lstrIdSer As String, lstrFiltro As String
        lstrIdAno = astrKeySer.Split(",")(0)
        lstrIdSer = astrKeySer.Split(",")(1)
        lstrFiltro = ClsIdCliente_ItemProgramaFactDbl.SstrNombreCampoBd & " = " &
                adblIdCliente.ToString() & " AND " &
                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " = 0 " &
                " AND " & ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd &
                " = " & lstrIdSer
        Return lstrFiltro
    End Function
    Private Shared Function FobjNuevaFact(aobjCliente As ClsCliente, astrIdPredAgru As String,
            adtmFechaFact As Date) As ClsFactura
        Dim lobjFactura = aobjCliente.ObjNuevaFactura(EnuModoFacturacionDef.EnuSistema)
        With lobjFactura
            .ObjEsPreFacturaBln.ObjValorPro = True
            .ObjIdPredioAgrupador_FacStr.ObjValorPro = astrIdPredAgru
            .ObjFechaFacturaDtm.ObjValorPro = adtmFechaFact
            ' Todas las facturas programadas la forma de pago es a crédito
            .ObjIdFormaPagoByt.ObjValorPro = EnuFormaPago.EnuCredito
            .ObjIdMedioPagoByt.ObjValorPro = EnuTipoMedioPagoDef.None
        End With
        Return lobjFactura
    End Function
    Private Shared Function FcolVlrItemPropietarios(aobjPredioAgr As ClsPredio,
            adecValorItem As Decimal, ablnAPropietaro As Boolean) As Collection
        Dim i = 0, lcolVlrItemProp As New Collection, ldblPorPart As Double
        Dim ldecVlrItemProp As Decimal
        Dim ldecVlrTotal = 0D
        If aobjPredioAgr.ObjIdTipoDestinatarioFacturaByt.ObjValorPro =
                EnuDestinatarioFacturaDef.EnuPropietario OrElse ablnAPropietaro Then
            Dim lcolPropropietarios As Collection = aobjPredioAgr.ColPropietarios
            For Each lobjProp As ClsPropietario In lcolPropropietarios
                ldblPorPart = lobjProp.ObjPorcentajePartiDbl.ObjValorPro
                ldecVlrItemProp = FdecValorRedondeado(adecValorItem * ldblPorPart)
                If i < lcolPropropietarios.Count - 1 Then
                    lcolVlrItemProp.Add(ldecVlrItemProp,
                            lobjProp.ObjIdCliente_PropDbl.ToString())
                    ldecVlrTotal += ldecVlrItemProp
                Else
                    If ldecVlrItemProp + ldecVlrTotal <> adecValorItem Then
                        ldecVlrItemProp = adecValorItem - ldecVlrTotal
                    End If
                    lcolVlrItemProp.Add(ldecVlrItemProp,
                            lobjProp.ObjIdCliente_PropDbl.ToString())
                End If
                i += 1
            Next
        Else
            Dim lobjCLiente = aobjPredioAgr.ObjArrendatario
            ldecVlrItemProp = adecValorItem
            lcolVlrItemProp.Add(ldecVlrItemProp, lobjCLiente.ObjIdClienteDbl.ToString())
        End If
        Return lcolVlrItemProp
    End Function
#End Region
#Region "Reversar Prefacturas"
    Friend Sub SReversePreFacturas()
        Dim lblnNoHayError = False
        GobjPanDat.SControleProcesoObj(True)
        Try
            GobjPanDat.SInicialiceTransaccion()
            If FblnHayPrefacturas() Then
                MobjArgumentoEventoPan = Nothing
                ObjArgumentoEventoPan.EnuProceso = EnuProcesoDef.enuRevPreFac
                ObjArgumentoEventoPan.DblCantAProcesar = 2
                RaiseEvent EvnInicio(Me, ObjArgumentoEventoPan)
                ObjArgumentoEventoPan.DblCantProcesada = 0
                RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
                SReverseItemsProgramaFact()
                ObjArgumentoEventoPan.DblCantProcesada = 1
                RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
                SElimineRegistrosPreFacturas()
                ObjArgumentoEventoPan.DblCantProcesada = 2
                RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
                GobjPanorama.SRegistreAccionLogApp("clsOrionCop", "Reversa Prefacturas")
            End If
            lblnNoHayError = True
        Catch ex As ErrorInesperadoPanDatException
            Throw
        Catch ex As ErrorInesperadoPanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As PanLException
            Throw
        Catch ex As ArgumentException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                GobjPanDat.SConfirmeTransaccion()
                GobjPanDat.SControleProcesoObj(False)
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    Private Shared Sub SReverseItemsProgramaFact()
        SReverseItemsProgramaFactPredios()
        SReverseItemsProgramaFactClientes()
    End Sub
    Private Shared Sub SReverseItemsProgramaFactPredios()
        Dim lstrSql = "SELECT DISTINCT P." & ClsIdPredioStr.SstrNombreCampoBd & ", " &
                ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd & "," &
                ClsIdServicio_ItemFactShr.SstrNombreCampoBd & " FROM (SELECT " &
                StrCampoCarpeta & ", " & ClsIdCentroUtilShr.
                SstrNombreCampoBd & ", " & ClsIdPredioStr.SstrNombreCampoBd & ", " &
                ClsIdPredioAgrupadorStr.SstrNombreCampoBd & " FROM " &
                ClsPredio.SstrNombreTabla & " WHERE " & ClsIdCarpetaShr.
                SstrNombreCampoBd & " = " & GshrIdCarpeta & " AND " & ClsIdCentroUtilShr.
                SstrNombreCampoBd & " = " & GshrIdCentroUtil & ") As P INNER JOIN " &
                ClsItemFactura.SstrNombreTabla & " AS I ON P." &
                StrCampoCarpeta & " = I." & ClsIdCarpetaShr.
                SstrNombreCampoBd & " AND P." & StrCampoCentroUtil &
                " = I." & StrCampoCentroUtil & " AND P." &
                ClsIdPredioAgrupadorStr.SstrNombreCampoBd & " = I." &
                ClsIdPredio_ItemFactStr.SstrNombreCampoBd & " WHERE " &
                ClsPrefijo_ItemFactStr.SstrNombreCampoBd & " = '" & GCSTRPREFPREFACTURA &
                "' ORDER BY " & ClsIdPredioStr.SstrNombreCampoBd
        Dim ldtbResul = ClsPanorama.FdtbDataTable(lstrSql)
        Dim lstrIdPredio As String, lshrIdAno As Short, lshrIdSer As Short
        For Each ldrwRes As DataRow In ldtbResul.Rows
            lstrIdPredio = ClsPanorama.FobjValorCampo(ldrwRes(
                    ClsIdPredioStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            lshrIdAno = ClsPanorama.FobjValorCampo(ldrwRes(
                    ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            lshrIdSer = ClsPanorama.FobjValorCampo(ldrwRes(
                    ClsIdServicio_ItemFactShr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            SReverseItemProgFact(0, lstrIdPredio, lshrIdAno, lshrIdSer)
        Next
    End Sub
    Private Shared Sub SReverseItemsProgramaFactClientes()
        Dim lstrTablaPri = ClsItemFactura.SstrNombreTabla
        Dim lstrTablaSec = ClsFactura.SstrNombreTabla
        Dim lstrCampSelPri As String() = {"DISTINCT " &
                ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd,
                ClsIdServicio_ItemFactShr.SstrNombreCampoBd}
        Dim lstrCampSelSec As String() = {ClsIdCliente_FactDbl.SstrNombreCampoBd}
        Dim lstrCampRelPri As String() = {StrCampoCarpeta,
                StrCampoCentroUtil,
                ClsIdFactura_ItemFactEnt.SstrNombreCampoBd}
        Dim lstrCampRelSec As String() = {StrCampoCarpeta,
                StrCampoCentroUtil, ClsIdFacturaEnt.SstrNombreCampoBd}
        Dim lstrFiltro = StrFiltroUbicacion_Pri & " And P." &
                ClsPrefijo_ItemFactStr.SstrNombreCampoBd & " = '" &
                GCSTRPREFPREFACTURA & "'"
        Dim lstrOrden As String(,) = {{ClsIdCliente_FactDbl.SstrNombreCampoBd, "ASC"},
                {ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd, "DESC"},
                {ClsIdServicio_ItemFactShr.SstrNombreCampoBd, "ASC"}}
        Dim lstrGrupo As String() = {}
        Dim ldtbResul = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCampSelPri,
                lstrTablaSec, lstrCampSelSec, lstrCampRelPri, lstrCampRelSec,
                lstrOrden, True, lstrFiltro, lstrGrupo)
        Dim ldblIdCliente As Double, lshrIdAno As Short, lshrIdSer As Short
        For Each ldrwRes As DataRow In ldtbResul.Rows
            ldblIdCliente = ClsPanorama.FobjValorCampo(ldrwRes(
                    ClsIdCliente_FactDbl.SstrNombreCampoBd),
                    EnuTipoValor.enuDouble)
            lshrIdAno = ClsPanorama.FobjValorCampo(ldrwRes(
                    ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            lshrIdSer = ClsPanorama.FobjValorCampo(ldrwRes(
                    ClsIdServicio_ItemFactShr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            SReverseItemProgFact(ldblIdCliente, "", lshrIdAno, lshrIdSer)
        Next
    End Sub
    Friend Shared Sub SReverseItemProgFact(adblIdCliente As Double,
            astrIdPredio As String, ashrIdAno As Short, ashrIdServicio As Short)
        Dim lstrTabla = ClsItemProgramaFact.SstrNombreTabla
        Dim lstrSql = "UPDATE " & lstrTabla & " SET " &
                ClsSaldo_ItemProgramaFactDec.SstrNombreCampoBd & " = " &
                ClsSaldo_ItemProgramaFactDec.SstrNombreCampoBd & " + " &
                ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd
        Dim lstrFiltro = " WHERE " & StrFiltroUbicacion & " AND " &
                ClsIdCliente_ItemProgramaFactDbl.SstrNombreCampoBd & " = " &
                adblIdCliente & " AND " &
                ClsIdPredio_ItemProgramaFactStr.SstrNombreCampoBd & " = '" &
                astrIdPredio & "' AND " & ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd &
                " = " & ashrIdAno & " AND " &
                ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd & " = " &
                ashrIdServicio
        lstrSql &= lstrFiltro
        Dim lentRes = GobjPanDat.SEjecuteSentenciaSql(lstrSql)
    End Sub
    Private Shared Sub SElimineRegistrosPreFacturas()
        Dim lcolNombresCamposRef = New Collection
        Dim lcolDatosRef = New Collection
        lcolNombresCamposRef.Add(StrCampoCarpeta)
        lcolNombresCamposRef.Add(StrCampoCentroUtil)
        lcolNombresCamposRef.Add(ClsEsPreFacturaBln.SstrNombreCampoBd)
        lcolDatosRef.Add(GshrIdCarpeta)
        lcolDatosRef.Add(GshrIdCentroUtil)
        lcolDatosRef.Add(True)
        GobjPanDat.SElimineRegistro(ClsFactura.SstrNombreTabla, lcolNombresCamposRef, lcolDatosRef)
        GobjPanDat.SElimineRegistro(ClsItemFactura.SstrNombreTabla, lcolNombresCamposRef, lcolDatosRef)
        GobjPanDat.SElimineRegistro(ClsNovedad.SstrNombreTabla, lcolNombresCamposRef, lcolDatosRef)
    End Sub
#End Region
#End Region
#Region "Generar Facturas Definitivas"
    Friend Function FblnGeneroFacturasDef(ByRef astrMens As String,
            ByRef ablnFacturoCompleto As Boolean) As Boolean
        Dim lblnNoHayError = False, lblnGenFact = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            GobjPanDat.SInicialiceTransaccion()
            MobjArgumentoEventoPan = Nothing
            BlnGenFacturas = True
            Dim ldtmFechaEstado = FdtmFechaPrefacturas.AddDays(-1)
            Dim ldtmFechaUltEst As Date = GobjParametros.ObjAnoActual.
                    ObjFechaUltEstadoCtaDtm.ObjValorPro
            If Not GobjParametros.ObjAnoActual.FblnFacturacionGenerada Then
                If ldtmFechaEstado > ldtmFechaUltEst.AddDays(-1) Then
                    SGenereEstadosCuentaClientes(ldtmFechaEstado)
                    If Not ObjArgumentoEventoPan.BlnCancele Then
                        SGenereEstadoCtaPrediosAgr(ldtmFechaEstado)
                    End If
                    GobjParametros.ObjAnoActual.SActualiceFechaGenEstados(ldtmFechaEstado)
                Else
                    SGenereEstadosCtaPrefacturas()
                End If
            Else
                If GobjParametros.ObjFechaUltCausacionGralDtm.ObjValorPro < Date.Today Then
                    SCauseMoraPrefacturas(astrMens)
                End If
                SGenereEstadosCtaPrefacturas()
            End If
            If Not ObjArgumentoEventoPan.BlnCancele Then
                SGenereFacturas(ablnFacturoCompleto)
            End If
            If Not ObjArgumentoEventoPan.BlnCancele Then
                If Not GobjParametros.ObjAnoActual.ObjPeriodoActual.BlnPeriodoFacturado Then
                    GobjParametros.SRegistreFechaFacturación()
                End If
                GobjPanorama.SRegistreAccionLogApp("clsOrionCop", "Genera Facturas Def.")
            End If
            If Not ablnFacturoCompleto Then
                SReverseEstadosCuenta()
                SReversePreFacturas()
            End If
            lblnNoHayError = True
        Catch ex As ErrorInesperadoPanDatException
            Throw
        Catch ex As ErrorInesperadoPanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As PanLException
            Throw
        Catch ex As ArgumentException
            Throw
        Catch ex As Exception
            Throw
        Finally
            BlnGenFacturas = False
            If Not lblnNoHayError OrElse ObjArgumentoEventoPan.BlnCancele Then
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
                If ObjArgumentoEventoPan.BlnCancele Then
                    ObjArgumentoEventoPan.EnuProceso = EnuProcesoDef.None
                    ObjArgumentoEventoPan.DblCantAProcesar = 0
                    RaiseEvent EvnInicio(Me, ObjArgumentoEventoPan)
                End If
            Else
                GobjPanDat.SConfirmeTransaccion()
                GobjPanDat.SControleProcesoObj(False)
                lblnGenFact = True
            End If
        End Try
        Return lblnGenFact
    End Function
    Private Sub SGenereFacturas(ByRef ablnFacturoCompleto As Boolean)
        Dim lentIdUltFacHabilitada = 0
        If GobjParametros.BlnEFacAutorizado Then
            lentIdUltFacHabilitada = GobjParametros.ObjRangoFraFinEnt.ObjValorPro
        End If
        BlnFacturando = True
        ObjArgumentoEventoPan.BlnCancele = False
        SModifiqueADefinitivas(lentIdUltFacHabilitada, ablnFacturoCompleto)
        BlnFacturando = False
        If Not ObjArgumentoEventoPan.BlnCancele Then
            SApliqueAnticipos()
            SActualiceEstadosFactAuto()
            DtmFechaFactAuto = GCDTMFECHANULA
            If GobjParametros.ObjServicioIdActivoBln.ObjValorPro Then
                SAdicioneServicioId()
            End If
        End If
        BlnFacturando = False
    End Sub
    Private Sub SModifiqueADefinitivas(aentIdUltFacHabilitada As Integer,
            ByRef ablnFactCompleto As Boolean)
        Dim i = 0.0
        Dim ldtbPrefacturas As DataTable = FdtbPrefacturas()
        If ldtbPrefacturas.Rows.Count > 0 Then
            Dim ldblCantAprocesar = FdblCantidadPreFacAProcesar(ldtbPrefacturas)
            ObjArgumentoEventoPan.EnuProceso = EnuProcesoDef.enuPasandoAFact
            ObjArgumentoEventoPan.BlnCancele = False
            ObjArgumentoEventoPan.DblCantAProcesar = ldblCantAprocesar
            RaiseEvent EvnInicio(Me, ObjArgumentoEventoPan)
            Dim lstrPrefijoPrefac = GCSTRPREFPREFACTURA
            Dim lobjFactura As New ClsFactura()
            ablnFactCompleto = True
            For Each ldrwPreFact As DataRow In ldtbPrefacturas.Rows
                i += 1
                Dim lentIdFactura As Integer = ClsPanorama.FobjValorCampo(
                        ldrwPreFact(ClsIdFacturaEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
                If aentIdUltFacHabilitada = 0 OrElse lentIdFactura <= aentIdUltFacHabilitada Then
                    lobjFactura.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefijoPrefac,
                            lentIdFactura})
                    If DtmFechaFactAuto = GCDTMFECHANULA Then
                        DtmFechaFactAuto = lobjFactura.ObjFechaFacturaDtm.ObjValorPro
                    Else
                        If lobjFactura.ObjFechaFacturaDtm.ObjValorPro <> DtmFechaFactAuto Then
                            DtmFechaFactAuto = lobjFactura.ObjFechaFacturaDtm.ObjValorPro
                        End If
                    End If
                    lobjFactura.SModifiqueADefinitiva()
                    Dim lstrPrefijoFac = GobjParametros.FstrPrefijoDoc(
                            EnuTipoDocOri.EnuFactura)
                    lobjFactura.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefijoFac,
                            lentIdFactura})
                    SEnlaceConEstadoCuenta(lobjFactura)
                    ObjArgumentoEventoPan.DblCantProcesada = i
                    RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
                    If ObjArgumentoEventoPan.BlnCancele Then
                        Exit Sub
                    End If
                Else
                    ablnFactCompleto = False
                    Exit For
                End If
            Next
        End If
    End Sub
    Private Shared Sub SEnlaceConEstadoCuenta(aobjFactura As ClsFactura)
        Dim lobjEstadoCuenta As ClsEstadoCuenta = FobjEstadoCuenta(aobjFactura)
        If IsNothing(lobjEstadoCuenta) Then
            SAdicioneEstadoCuenta(aobjFactura)
        Else
            Dim ldtmFechaEstado As Date = aobjFactura.ObjFechaFacturaDtm.ObjValorPro
            ldtmFechaEstado = ldtmFechaEstado.AddDays(-1)
            If Not CType(lobjEstadoCuenta.EnuPermisosObj And EnuPermisosDef.enuModificar, Boolean) Then
                lobjEstadoCuenta.EnuPermisosObj += EnuPermisosDef.enuModificar
            End If
            With lobjEstadoCuenta
                .EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                .ObjFechaEstadoDtm.ObjValorPro = ldtmFechaEstado
                .ObjPrefijoFac_EstadoStr.ObjValorPro = aobjFactura.ObjPrefijo_FactStr.ObjValorPro
                .ObjIdFactura_EstadoEnt.ObjValorPro = aobjFactura.ObjIdFacturaEnt.ObjValorPro
                .SActualice(True)
            End With
        End If
    End Sub
    Private Shared Sub SAdicioneEstadoCuenta(aobjFactura As ClsFactura)
        Dim ldtmFecha As Date = aobjFactura.ObjFechaFacturaDtm.ObjValorPro
        ldtmFecha = ldtmFecha.AddDays(-1)
        Dim ldblIdCliente As Double = aobjFactura.ObjIdCliente_FactDbl.ObjValorPro
        Dim lstrIdPredioAgr As String = aobjFactura.ObjIdPredioAgrupador_FacStr.ObjValorPro
        Dim lstrKeySerUnico = String.Empty
        Dim lblnFactPorServ = aobjFactura.BlnFacturaPorServicio
        If lblnFactPorServ Then
            lstrKeySerUnico = aobjFactura.StrKeySerUnico
        End If
        Dim lobjEstadoCuenta = FobjNuevoEstadoCuenta(ldblIdCliente, lstrIdPredioAgr, ldtmFecha)
        With lobjEstadoCuenta
            .ObjPrefijoFac_EstadoStr.ObjValorPro = aobjFactura.ObjPrefijo_FactStr.ObjValorPro
            .ObjIdFactura_EstadoEnt.ObjValorPro = aobjFactura.ObjIdFacturaEnt.ObjValorPro
            .ObjDeudaCapitalDec.ObjValorPro = 0
            .ObjDeudaIntMoraDec.ObjValorPro = 0
            .SActualice(True)
        End With
    End Sub
    Private Shared Sub SReverseEstadosCuenta()
        Dim lstrTabla = ClsEstadoCuenta.SstrNombreTabla
        Dim lstrFiltro = StrFiltroUbicacion & " AND " & ClsPrefijoFac_EstadoStr.SstrNombreCampoBd &
                " = '" & GCSTRPREFPREFACTURA & "'"
        Dim lstrSqlElim = ClsPanoramaDat.FstrConstruyaExpSqlEliminar(lstrTabla, lstrFiltro)
        GobjPanDat.SEjecuteSentenciaSql(lstrSqlElim)
    End Sub
    Private Shared Sub SAdicioneServicioId()
        Dim lobjFactura As New ClsFactura()
        Dim lobjServicioId As ClsServicio = GobjParametros.FobjServicioId
        Dim ldecValorSerId = 0D
        Dim lstrIdPredioAgr = String.Empty
        Dim lentValorMaxSerId = FentMaxValorSerId()
        Dim lstrUltFac = FstrIdUltimasFras(False)
        If Not String.IsNullOrEmpty(lstrUltFac) Then
            Dim lstrPrefijo = lstrUltFac.Split(";")(0)
            Dim lentIdFacIni = CType(lstrUltFac.Split(";")(1), Integer)
            Dim lentIdFacFin = CType(lstrUltFac.Split(";")(2), Integer)
            For lentIdFactura As Integer = lentIdFacIni To lentIdFacFin
                lobjFactura.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefijo, lentIdFactura})
                lstrIdPredioAgr = lobjFactura.ObjIdPredioAgrupador_FacStr.ToString
                If Not String.IsNullOrEmpty(lstrIdPredioAgr) Then
                    ldecValorSerId = FdecValorSerIdDefinitivo(lobjFactura, lentValorMaxSerId)
                    If ldecValorSerId > 0 Then
                        lobjFactura.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                        Dim lobjNuevoIetmFac = lobjFactura.FobjNuevoItemFactura
                        Dim lobjItemFac As ClsItemFactura = lobjFactura.ColItemsFactura(1)
                        Dim lobjSerAdmin As ClsServicio = lobjItemFac.ObjServicio
                        With lobjNuevoIetmFac
                            .ObjFechaGraciaIFDtm.ObjValorPro = lobjSerAdmin.DtmFechaGraciaPeriActual
                            .ObjFechaVencimientoIFDtm.ObjValorPro =
                                    lobjSerAdmin.DtmFechaVencePeriActual
                            .ObjPrefijo_ItemFactStr.ObjValorPro =
                                lobjFactura.ObjPrefijo_FactStr.ObjValorPro
                            .ObjIdFactura_ItemFactEnt.ObjValorPro =
                                lobjFactura.ObjIdFacturaEnt.ObjValorPro
                            .ObjDebitos_ItemFactDec.ObjValorPro = ldecValorSerId
                            .ObjIdAno_ServicioItemFactShr.ObjValorPro =
                                    lobjServicioId.ObjIdAno_ServicioShr.ObjValorPro
                            .ObjIdServicio_ItemFactShr.ObjValorPro =
                                    lobjServicioId.ObjIdServicioShr.ObjValorPro
                            .ObjDetalle_ItemFactStr.ObjValorPro =
                                    lobjServicioId.ObjNombreServicioStr.ObjValorPro
                            .ObjIdPredio_ItemFactStr.ObjValorPro =
                                    lobjFactura.ObjIdPredioAgrupador_FacStr.ObjValorPro
                            .ObjEsPrefactura_ItemFactBln.ObjValorPro =
                                    lobjFactura.ObjEsPreFacturaBln.ObjValorPro
                            .ObjValor_ItemFactDec.ObjValorPro = ldecValorSerId
                        End With
                        lobjFactura.SAdicioneItemSerId(lobjNuevoIetmFac)
                        lobjFactura.SActualice(True)
                    End If
                End If
            Next
        End If
    End Sub
    ''' <summary>
    ''' Devuelve el valor del servicio de identificación que se debe cobrar para que 
    ''' la factura termine exactamente en el valor del servicio definido para el 
    ''' predio agrupador
    ''' </summary>
    Private Shared Function FdecValorSerIdDefinitivo(aobjFactura As ClsFactura,
            aentValorMaxSerId As Integer) As Decimal
        If aobjFactura.DecDeuda = 0 OrElse aobjFactura.FblnSerIdAplicado Then
            Return 0
        End If
        Dim lobjPredioAgr As New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
        Dim lobjEstadoCta As ClsEstadoCuenta
        Dim ldecSaldoFact As Decimal, ldecTotAPagar As Decimal, ldecValorSerId As Decimal
        Dim ldecResiduo As Decimal
        Dim lstrIdPredioAgr As String = aobjFactura.ObjIdPredioAgrupador_FacStr.ObjValorPro
        lobjPredioAgr.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrIdPredioAgr})
        lobjEstadoCta = aobjFactura.ObjEstadoCuenta
        ldecSaldoFact = aobjFactura.DecDeuda
        ldecTotAPagar = ldecSaldoFact + lobjEstadoCta.DecTotalDeuda
        ldecValorSerId = lobjPredioAgr.ObjValorServicioIdDec.ObjValorPro
        If ldecTotAPagar > 0 Then
            If aentValorMaxSerId = 100 Then
                ldecResiduo = ldecTotAPagar Mod 100
            ElseIf aentValorMaxSerId = 1000 Then
                ldecResiduo = ldecTotAPagar Mod 1000
            Else
                Throw New ErrorInesperadoPanLException("El valor del servicio de identificación para los Predios " &
            "no ha sido ingresado correctamente!")
            End If
            If ldecResiduo <= ldecValorSerId Then
                ldecValorSerId -= ldecResiduo
            Else
                If ldecResiduo < aentValorMaxSerId Then
                    ldecValorSerId = aentValorMaxSerId + ldecValorSerId - ldecResiduo
                End If
            End If
        Else
            ldecValorSerId = 0
        End If
        Return ldecValorSerId
    End Function
    Private Shared Function FentMaxValorSerId() As Integer
        Dim lstrTabla = ClsPredio.SstrNombreTabla
        Dim lstrCamposSelect = {"MAX(" & ClsValorServicioIdDec.SstrNombreCampoBd & ")"}
        Dim lstrFiltro = StrFiltroUbicacion
        Dim ldtbTabla = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamposSelect, {{"", ""}}, lstrFiltro)
        Dim lentValorMax = ClsPanorama.FobjValorCampo(ldtbTabla.Rows(0)(0), EnuTipoValor.enuInteger)
        If lentValorMax < 100 Then
            lentValorMax = 100
        ElseIf lentValorMax < 1000 Then
            lentValorMax = 1000
        Else
            lentValorMax = 0
        End If
        Return lentValorMax
    End Function
    ''' <summary>
    ''' Establece la relación entre el estado de cuenta y la factura pasada en el argumento.
    ''' </summary>
    Private Shared Function FobjEstadoCuenta(aobjFactura As ClsFactura) As ClsEstadoCuenta
        Dim lobjEstadoCuenta As ClsEstadoCuenta = Nothing
        Dim lentIdFact As Integer = aobjFactura.ObjIdFacturaEnt.ObjValorPro
        Dim lstrTabla = ClsEstadoCuenta.SstrNombreTabla
        Dim lstrCamposSelect = {"*"}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " &
                ClsPrefijoFac_EstadoStr.SstrNombreCampoBd & " = '" & GCSTRPREFPREFACTURA &
                "' AND " & ClsIdFactura_EstadoEnt.SstrNombreCampoBd & " = " &
                lentIdFact
        Dim lstrOrden(,) = {{ClsIdEstadoCuentaEnt.SstrNombreCampoBd, "DESC"}}
        Dim ldtbEstadosCuenta = ClsPanorama.FdtbDataTable(ClsEstadoCuenta.SstrNombreTabla,
                lstrCamposSelect, lstrOrden, lstrFiltro)
        If ldtbEstadosCuenta.Rows.Count >= 1 Then
            Dim ldrwEstadoCta = ldtbEstadosCuenta.Rows(0)
            lobjEstadoCuenta = New ClsEstadoCuenta(ldrwEstadoCta)
            lobjEstadoCuenta.SLeaValores(True)
        End If
        Return lobjEstadoCuenta
    End Function
    Private Sub SApliqueAnticipos()
        Dim i = 0.0
        Dim ldtbAnticiosPorAplicar = FdtbAnticiposPorAplicar()
        If ldtbAnticiosPorAplicar.Rows.Count > 0 Then
            ObjArgumentoEventoPan.BlnCancele = False
            ObjArgumentoEventoPan.DblCantAProcesar = ldtbAnticiosPorAplicar.Rows.Count
            ObjArgumentoEventoPan.EnuProceso = EnuProcesoDef.enuApliAnti
            RaiseEvent EvnInicio(Me, ObjArgumentoEventoPan)
            ObjArgumentoEventoPan.DblCantProcesada = i
            RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
            For Each ldrwAntPorApl As DataRow In ldtbAnticiosPorAplicar.Rows
                i += 1
                SApliqueAnticipo(ldrwAntPorApl)
                ObjArgumentoEventoPan.DblCantProcesada = i
                RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
                If ObjArgumentoEventoPan.BlnCancele Then
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Shared Sub SApliqueAnticipo(adrwAntPorApl As DataRow)
        Dim lobjAnticipo As New ClsAnticipo(EnuModoInstanciaObjDef.enuUnico)
        Dim lentIdAnticipo As Integer = ClsPanorama.FobjValorCampo(adrwAntPorApl(
                ClsIdAnticipoEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
        Dim lobjValorLlave() As Object = {GshrIdCarpeta, GshrIdCentroUtil, lentIdAnticipo}
        lobjAnticipo.SAbra(lobjValorLlave)
        Dim lenuOrigenAnt As EnuTipoDocOri = lobjAnticipo.ObjIdTipoDocOrigen_AntByt.ObjValorPro
        Select Case lenuOrigenAnt
            Case EnuTipoDocOri.EnuReciboCaja
                SApliqueAnticipoRecibo(lobjAnticipo)
            Case EnuTipoDocOri.EnuNotaAjuste
                SApliqueAnticipoAjuste(lobjAnticipo)
        End Select
    End Sub
    Private Shared Sub SApliqueAntFacMan(aobjFact As ClsFactura)
        Dim ldecVlrAplicado = 0D, ldecValorDeuda As Decimal = aobjFact.DecDeuda
        Dim lcolAnt As Collection = FcolAnticipo(aobjFact)
        If lcolAnt.Count > 0 Then
            Dim lstrServicios As String()
            Dim lstrIdPredioAgru As String
            Dim ldecVlrPorAplicar As Decimal
            Dim ldblIdCliente As Double
            Dim lobjCliente As New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
            Dim ldtmFecFac As Date = aobjFact.ObjFechaFacturaDtm.ObjValorPro
            For Each lobjAnt As ClsAnticipo In lcolAnt
                If ldecValorDeuda > 0 Then
                    ldecVlrAplicado = 0
                    lstrServicios = lobjAnt.ObjServicios_AntStr.ToString().Split(",")
                    lstrIdPredioAgru = lobjAnt.ObjIdPredioAgrupador_AntStr.ToString()
                    ldecVlrPorAplicar = lobjAnt.DecAnticipoPorAplicar
                    ldblIdCliente = lobjAnt.ObjIdCliente_AntDbl.ObjValorPro
                    lobjCliente.SAbra({GshrIdCarpeta, GshrIdCentroUtil, ldblIdCliente})
                    SApliqueAnticipoFac(aobjFact, lstrServicios, ldecVlrPorAplicar, lobjCliente,
                            ldtmFecFac, lstrIdPredioAgru, lobjAnt, ldecVlrAplicado)
                End If
                ldecValorDeuda -= ldecVlrAplicado
            Next
        End If
    End Sub
    Private Shared Sub SApliqueAnticipoRecibo(aobjAnticipo As ClsAnticipo)
        Dim lstrIdPredioAgru As String =
                aobjAnticipo.ObjIdPredioAgrupador_AntStr.ToString()
        Dim lstrServicios As String() =
                aobjAnticipo.ObjServicios_AntStr.ToString().Split(",")
        Dim ldblIdCliente As Double = aobjAnticipo.ObjIdCliente_AntDbl.ObjValorPro
        Dim ldecVlrPorAplicar As Decimal = aobjAnticipo.DecAnticipoPorAplicar
        Dim ldecVlrAplicado As Decimal = 0, lstrPref As String, lentIdFac As Integer
        Dim lobjCliente As New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
        lobjCliente.SAbra({GshrIdCarpeta, GshrIdCentroUtil, ldblIdCliente})
        Dim lobjFactura As New ClsFactura()
        Dim ldtbIdFactsFecha = FdtbIdFactVivasPredio(ldblIdCliente, lstrIdPredioAgru)
        For Each ldrwIdFac As DataRow In ldtbIdFactsFecha.Rows
            lstrPref = ClsPanorama.FobjValorCampo(ldrwIdFac(
                    ClsPrefijo_FactStr.SstrNombreCampoBd), EnuTipoValor.enuString)
            lentIdFac = ClsPanorama.FobjValorCampo(ldrwIdFac(
                    ClsIdFacturaEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
            lobjFactura.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentIdFac})
            If lobjFactura.ObjPrefijo_FactStr.ObjValorPro <> GCSTRPREFPREFACTURA Then
                ldecVlrPorAplicar -= ldecVlrAplicado
                If ldecVlrPorAplicar > 0 Then
                    SApliqueAnticipoFac(lobjFactura, lstrServicios, ldecVlrPorAplicar,
                            lobjCliente, DtmFechaFactAuto, lstrIdPredioAgru,
                            aobjAnticipo, ldecVlrAplicado)
                Else
                    Exit For
                End If
            End If
        Next
    End Sub
    Private Shared Sub SApliqueAnticipoFac(aobjFactura As ClsFactura, astrServicios As String(),
            adecVlrPorAplicar As Decimal, aobjCliente As ClsCliente, adtmFechaFact As Date,
            astrIdpreAgr As String, aobjAnticipo As ClsAnticipo, ByRef adecVlrAplicado As Decimal)
        Dim lobjNotaCon As ClsNotaCon = Nothing
        Dim ldecVlrPorAplicar As Decimal = adecVlrPorAplicar, ldecVlrAplicado As Decimal
        Dim ldecTotVlrApli As Decimal
        Dim lshrIdIemFac As Short
        Dim lstrPeriodoFact = ClsPanorama.FstrPeriodo(aobjFactura.ObjFechaFacturaDtm.ObjValorPro)
        Dim lblnAplicar As Boolean = aobjFactura.ObjPrefijo_FactStr.ObjValorPro <> GCSTRPREFPREFACTURA
        If lblnAplicar Then
            For Each lobjItemFact As ClsItemFactura In aobjFactura.ColItemsFactura
                If GobjParametros.ObjPermiteAnticipoPorServicioBln.ObjValorPro Then
                    lblnAplicar = astrServicios.Contains("A")
                    If Not lblnAplicar Then
                        lblnAplicar = astrServicios.Contains(lobjItemFact.StrServicio)
                    End If
                End If
                If lblnAplicar Then
                    lshrIdIemFac = lobjItemFact.ObjIdItemFacturaShr.ObjValorPro
                    If IsNothing(lobjNotaCon) Then
                        lobjNotaCon = aobjCliente.FobjNuevaNotaCon(adtmFechaFact, astrIdpreAgr)
                        lobjNotaCon.ObjIdAnticipo_NotaConEnt.ObjValorPro =
                                aobjAnticipo.ObjIdAnticipoEnt.ObjValorPro
                        lobjNotaCon.ObjPrefijo_NotaRCrStr.ObjValorPro = String.Empty
                        lobjNotaCon.ObjIdNotaRCrEnt.ObjValorPro = 0
                    End If
                    If ldecVlrPorAplicar > 0 Then
                        ldecVlrAplicado = 0
                        SGenereItemsNotaCon(lobjNotaCon, aobjFactura, ldecVlrPorAplicar,
                                lshrIdIemFac, ldecVlrAplicado)
                        ldecVlrPorAplicar -= ldecVlrAplicado
                        ldecTotVlrApli += ldecVlrAplicado
                        adecVlrAplicado += ldecVlrAplicado
                    End If
                End If
                If ldecVlrPorAplicar = 0 Then Exit For
            Next
            If Not IsNothing(lobjNotaCon) AndAlso ldecTotVlrApli > 0 Then
                lobjNotaCon.SActualice(True)
                SApliqueNotaCon(lobjNotaCon, aobjFactura)
                aobjAnticipo.SGenereNovedadAntAplicado(lobjNotaCon)
                aobjAnticipo.SActualice(True)
            End If
        End If
    End Sub
    ' Genera los Items de la Nota Contable para la aplicación de un aticipo generado en un
    ' recibo de caja
    Private Shared Sub SGenereItemsNotaCon(aobjNotaCon As ClsNotaCon, aobjFactura As ClsFactura,
            adecValorPorAplicar As Decimal, ashrIdItemFcat As Short,
            ByRef adecValorAplicado As Decimal)
        Dim ldecDsctoPP = 0D, ldecDeudaCapItemFac As Decimal
        Dim ldecVlrAAplicar As Decimal
        For Each lobjItemFac As ClsItemFactura In aobjFactura.ColItemsFactura
            If ashrIdItemFcat = 0 OrElse lobjItemFac.ObjIdItemFacturaShr.ObjValorPro = ashrIdItemFcat Then
                ldecVlrAAplicar = 0
                If Not lobjItemFac.FblnRetencionAplicada(EnuTipoDescuentoDef.EnuDsctoPP) Then
                    ldecDsctoPP = lobjItemFac.FdecDsctoPPAAplicar(aobjFactura.ObjFechaFacturaDtm.ObjValorPro)
                End If
                ldecDeudaCapItemFac = lobjItemFac.FdecDeudaServicioTotal - lobjItemFac.FdecDeudaIva
                If ldecDsctoPP > 0 AndAlso ldecDeudaCapItemFac > 0 Then
                    If adecValorPorAplicar >= ldecDeudaCapItemFac - ldecDsctoPP Then
                        ldecDeudaCapItemFac -= ldecDsctoPP
                        aobjNotaCon.SGenereItemNotaCon(ldecDsctoPP,
                                aobjFactura.ObjPrefijo_FactStr.ObjValorPro,
                                aobjFactura.ObjIdFacturaEnt.ObjValorPro,
                                EnuTipoItemNotaConDef.EnuDsctoPP,
                                lobjItemFac.ObjIdItemFacturaShr.ObjValorPro)
                    End If
                End If
                Dim ldecDeudaMora As Decimal = lobjItemFac.FdecDeudaIntMora
                If ldecDeudaMora > 0 Then
                    If adecValorPorAplicar >= ldecDeudaMora Then
                        ldecVlrAAplicar = ldecDeudaMora
                    Else
                        ldecVlrAAplicar = adecValorPorAplicar
                    End If
                    aobjNotaCon.SGenereItemNotaCon(ldecVlrAAplicar, aobjFactura.ObjPrefijo_FactStr.ObjValorPro,
                            aobjFactura.ObjIdFacturaEnt.ObjValorPro, EnuTipoItemNotaConDef.EnuAplicaAntInt,
                            lobjItemFac.ObjIdItemFacturaShr.ObjValorPro)
                    adecValorPorAplicar -= ldecVlrAAplicar
                    adecValorAplicado += ldecVlrAAplicar
                End If
                If adecValorPorAplicar >= ldecDeudaCapItemFac Then
                    ldecVlrAAplicar = ldecDeudaCapItemFac
                Else
                    ldecVlrAAplicar = adecValorPorAplicar
                End If
                If ldecVlrAAplicar > 0 Then
                    aobjNotaCon.SGenereItemNotaCon(ldecVlrAAplicar, aobjFactura.ObjPrefijo_FactStr.ObjValorPro,
                    aobjFactura.ObjIdFacturaEnt.ObjValorPro, EnuTipoItemNotaConDef.EnuAplicaAntCap,
                    lobjItemFac.ObjIdItemFacturaShr.ObjValorPro)
                End If
                adecValorPorAplicar -= ldecVlrAAplicar
                adecValorAplicado += ldecVlrAAplicar
            End If
        Next
    End Sub
    Private Shared Sub SApliqueAnticipoAjuste(aobjAnticipo As ClsAnticipo)
        Dim ldecDeudaFac As Decimal
        Dim lstrIdPredioAgru As String = aobjAnticipo.ObjIdPredioAgrupador_AntStr.ToString()
        Dim lstrIdPredio As String = aobjAnticipo.ObjIdPredio_AntStr.ObjValorPro
        Dim ldblIdCliente As Double = aobjAnticipo.ObjIdCliente_AntDbl.ObjValorPro
        Dim lstrServicios As String() = aobjAnticipo.ObjServicios_AntStr.ToString().Split(",")
        Dim ldecVlrPorAplicar As Decimal = aobjAnticipo.ObjValor_AntDec.ObjValorPro -
                aobjAnticipo.ObjDebitos_AntDec.ObjValorPro
        Dim lobjCliente As New ClsCliente(EnuModoInstanciaObjDef.EnuUnico)
        lobjCliente.SAbra({GshrIdCarpeta, GshrIdCentroUtil, ldblIdCliente})
        Dim lobjNotaCon As ClsNotaCon = Nothing
        Dim lcolFrasVivas = lobjCliente.FcolFacturasAuto(lstrIdPredioAgru, DtmFechaFactAuto)
        If lcolFrasVivas.Count > 0 AndAlso ldecVlrPorAplicar > 0 Then
            For Each lobjFactura As ClsFactura In lcolFrasVivas
                If lobjFactura.ObjFechaFacturaDtm.ObjValorPro = DtmFechaFactAuto AndAlso
                        ldecVlrPorAplicar > 0 Then
                    ldecDeudaFac = lobjFactura.DecDeuda
                    If ldecDeudaFac > 0 AndAlso
                            lobjFactura.ObjPrefijo_FactStr.ObjValorPro <>
                            GCSTRPREFPREFACTURA Then
                        If IsNothing(lobjNotaCon) Then
                            lobjNotaCon = lobjCliente.FobjNuevaNotaCon(DtmFechaFactAuto,
                                    lstrIdPredioAgru)
                            lobjNotaCon.ObjIdAnticipo_NotaConEnt.ObjValorPro =
                                    aobjAnticipo.ObjIdAnticipoEnt.ObjValorPro
                            lobjNotaCon.ObjPrefijo_NotaRCrStr.ObjValorPro = String.Empty
                            lobjNotaCon.ObjIdNotaRCrEnt.ObjValorPro = 0
                        End If
                        Dim ldecVlrAAplicar As Decimal
                        If ldecDeudaFac > ldecVlrPorAplicar Then
                            ldecVlrAAplicar = ldecVlrPorAplicar
                        Else
                            ldecVlrAAplicar = ldecDeudaFac
                        End If
                        SGenereItemsNotaCon(lobjNotaCon, lobjFactura, ldecVlrAAplicar,
                                lstrServicios)
                        If Not IsNothing(lobjNotaCon) AndAlso
                                lobjNotaCon.ColItemsNotaCon.Count > 0 Then
                            lobjNotaCon.SActualice(True)
                            SApliqueNotaCon(lobjNotaCon, lobjFactura)
                            aobjAnticipo.SGenereNovedadAntAplicado(lobjNotaCon)
                            aobjAnticipo.SActualice(True)
                            ldecVlrPorAplicar -= ldecVlrAAplicar
                        End If
                    End If
                End If
                If ldecVlrPorAplicar = 0 Then Exit For
            Next
        End If
    End Sub
    ' Genera los Items de la Nota Contable para la aplicación de un aticipo generado en ajuste
    ' Cuotas admin
    Private Shared Sub SGenereItemsNotaCon(aobjNotaCon As ClsNotaCon, aobjFactura As ClsFactura,
            adecValorPorAplicar As Decimal, astrServicios As String())
        Dim ldecDsctoPP As Decimal, ldecDeudaCapItemFac As Decimal
        Dim ldecVlrAAplicar As Decimal
        For Each lobjItemFac As ClsItemFactura In aobjFactura.ColItemsFactura
            If astrServicios.Contains("A") OrElse
                    astrServicios.Contains(lobjItemFac.StrServicio) Then
                ldecVlrAAplicar = 0
                If Not aobjNotaCon.BlnEsAjusteCuotaAdmin Then
                    ldecDsctoPP = lobjItemFac.FdecDsctoPPAAplicar(
                        aobjFactura.ObjFechaFacturaDtm.ObjValorPro)
                End If
                ldecDeudaCapItemFac = lobjItemFac.FdecDeudaServicioTotal - lobjItemFac.FdecDeudaIva
                If ldecDsctoPP > 0 AndAlso ldecDeudaCapItemFac > 0 Then
                    ldecDeudaCapItemFac -= ldecDsctoPP
                    aobjNotaCon.SGenereItemNotaCon(ldecDsctoPP,
                            aobjFactura.ObjPrefijo_FactStr.ObjValorPro,
                            aobjFactura.ObjIdFacturaEnt.ObjValorPro,
                            EnuTipoItemNotaConDef.EnuDsctoPP,
                            lobjItemFac.ObjIdItemFacturaShr.ObjValorPro)
                End If
                Dim ldecDeudaMora As Decimal = lobjItemFac.FdecDeudaIntMora
                If ldecDeudaMora > 0 Then
                    If adecValorPorAplicar >= ldecDeudaMora Then
                        ldecVlrAAplicar = ldecDeudaMora
                    Else
                        ldecVlrAAplicar = adecValorPorAplicar
                    End If
                    aobjNotaCon.SGenereItemNotaCon(ldecVlrAAplicar,
                            aobjFactura.ObjPrefijo_FactStr.ObjValorPro,
                            aobjFactura.ObjIdFacturaEnt.ObjValorPro,
                            EnuTipoItemNotaConDef.EnuAplicaAntInt,
                            lobjItemFac.ObjIdItemFacturaShr.ObjValorPro)
                    adecValorPorAplicar -= ldecVlrAAplicar
                End If
                If adecValorPorAplicar >= ldecDeudaCapItemFac Then
                    ldecVlrAAplicar = ldecDeudaCapItemFac
                Else
                    ldecVlrAAplicar = adecValorPorAplicar
                End If
                If ldecVlrAAplicar > 0 Then
                    aobjNotaCon.SGenereItemNotaCon(ldecVlrAAplicar,
                            aobjFactura.ObjPrefijo_FactStr.ObjValorPro,
                            aobjFactura.ObjIdFacturaEnt.ObjValorPro,
                            EnuTipoItemNotaConDef.EnuAplicaAntCap,
                            lobjItemFac.ObjIdItemFacturaShr.ObjValorPro)
                End If
                adecValorPorAplicar -= ldecVlrAAplicar
                If adecValorPorAplicar = 0 Then Exit For
            End If
        Next
    End Sub
    Private Shared Sub SApliqueNotaCon(aobjNotaCon As ClsNotaCon, aobjFactura As ClsFactura)
        For Each lobjItemNotaCon As ClsItemNotaCon In aobjNotaCon.ColItemsNotaCon
            Dim lstrPrefijoFra As String = lobjItemNotaCon.ObjPrefijoFact_ItemNotaConStr.ObjValorPro
            Dim lentIdFactura As Integer = lobjItemNotaCon.ObjIdFactura_ItemNotaConEnt.ObjValorPro
            Dim lstrNroFac As String = ClsPanorama.FstrNumeroDcto(lstrPrefijoFra, lentIdFactura)
            aobjFactura.SApliqueAnticipo(lobjItemNotaCon)
            aobjFactura.SActualice(True)
        Next
    End Sub
    Private Shared Function FdtbAnticiposPorAplicar() As DataTable
        Dim lstrNombreTabla = ClsAnticipo.SstrNombreTabla
        Dim lstrCamposSelect As String() = {ClsIdAnticipoEnt.SstrNombreCampoBd}
        Dim lstrOrden As String(,) = {{ClsIdAnticipoEnt.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " & ClsCreditos_AntDec.SstrNombreCampoBd &
                " <> " & ClsDebitos_AntDec.SstrNombreCampoBd
        Dim ldtbAntiPorApli As DataTable = ClsPanorama.FdtbDataTable(lstrNombreTabla,
                lstrCamposSelect, lstrOrden, lstrFiltro)
        Return ldtbAntiPorApli
    End Function
    Private Shared Function FdtbAnticiposPorAplicar(astrIdPreAgr As String,
            adblIdCliente As Double) As DataTable
        Dim lstrNombreTabla = ClsAnticipo.SstrNombreTabla
        Dim lstrCamposSelect As String() = {ClsIdAnticipoEnt.SstrNombreCampoBd}
        Dim lstrIndice As String(,) = {{ClsIdAnticipoEnt.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " & ClsCreditos_AntDec.SstrNombreCampoBd &
                " <> " & ClsDebitos_AntDec.SstrNombreCampoBd
        If Not String.IsNullOrEmpty(astrIdPreAgr) Then
            lstrFiltro &= " AND " & ClsIdPredioAgrupador_AntStr.SstrNombreCampoBd & " = '" &
                    astrIdPreAgr & "'"
        Else
            lstrFiltro &= " AND " & ClsIdCliente_AntDbl.SstrNombreCampoBd & " = " & adblIdCliente
        End If
        Dim ldtbAntiPorApli As DataTable = ClsPanorama.FdtbDataTable(lstrNombreTabla,
                lstrCamposSelect, lstrIndice, lstrFiltro)
        Return ldtbAntiPorApli
    End Function
    ''' <summary>
    ''' Devuelve un datatable con toda la información de los anticipos por aplicar
    ''' </summary>
    ''' <returns></returns>
    Private Shared Function FdtbAnticiposSinAplicar()
        Dim lstrTabla = ClsAnticipo.SstrNombreTabla
        Dim lstrCamposSelect As String() = {"*"}
        Dim lstrOrden As String(,) = {{ClsIdCliente_AntDbl.SstrNombreCampoBd, "ASC"},
                {ClsIdPredioAgrupador_AntStr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " & ClsCreditos_AntDec.SstrNombreCampoBd &
                " <> " & ClsDebitos_AntDec.SstrNombreCampoBd
        Dim ldtbAntPorApl = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamposSelect, lstrOrden,
                lstrFiltro, False, Array.Empty(Of String))
        Return ldtbAntPorApl
    End Function
    Private Shared Function FdtbIdFactVivasPredio(adblIdCliente As Double,
            astrIdPredAgru As String) As DataTable
        Dim lstrTabla = ClsFactura.SstrNombreTabla
        Dim lstrCamSel = {ClsPrefijo_FactStr.SstrNombreCampoBd, ClsIdFacturaEnt.SstrNombreCampoBd}
        Dim lstrOrden = {{ClsPrefijo_FactStr.SstrNombreCampoBd, "ASC"},
                {ClsIdFacturaEnt.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " & ClsIdCliente_FactDbl.SstrNombreCampoBd &
                " = " & adblIdCliente & " AND " & ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd &
                " = '" & astrIdPredAgru & "' AND " & ClsPrefijo_FactStr.SstrNombreCampoBd & " <> '" &
                GCSTRPREFPREFACTURA & "'" & " AND " & ClsDebitos_FactDec.SstrNombreCampoBd & " > " &
                ClsCreditos_FactDec.SstrNombreCampoBd
        Dim ldtbIdFacsFechaPred = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamSel, lstrOrden,
                lstrFiltro)
        Return ldtbIdFacsFechaPred
    End Function
    ''' <summary>
    ''' devuelve un string que contiene elprefijo y los numeros de la primera y las ultimas notas con
    ''' (Aplicacion Anticipos) generadas, separadas por un punto y coma.
    ''' </summary>
    Friend Shared Function FstrIdUltimasNotasAplAnt() As String
        GobjPanDat.SControleProcesoObj(True)
        Dim lstrIdUltimasNotasAA = String.Empty
        Dim lstbSql As New StringBuilder
        Dim ldtmFecha = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
        Dim lstrFecha = "'" & ClsPanoramaDat.FstrFechaNormalizada(ldtmFecha) & "'"
        Dim ldrwIdNotasCon() As DataRow = Nothing
        If ldtmFecha <> GCDTMFECHANULA Then
            ' Campos a seleccionar
            Dim lstrCamposSelect = {ClsPrefijo_NotaConStr.SstrNombreCampoBd,
                    ClsIdNotaConEnt.SstrNombreCampoBd}
            ' Filtro
            With lstbSql
                .Clear()
                .Append(StrFiltroUbicacion).Append(" AND ")
                .Append(ClsFecha_NotaConDtm.SstrNombreCampoBd).Append(" >= ").Append(lstrFecha)
            End With
            Dim lstrFiltro = lstbSql.ToString
            Dim ldtbIdNotasCon = ClsPanorama.FdtbDataTable(ClsNotaCon.SstrNombreTabla, lstrCamposSelect,
                    {{ClsIdNotaConEnt.SstrNombreCampoBd, "ASC"}}, lstrFiltro)
            ldrwIdNotasCon = ldtbIdNotasCon.Select()
            Dim lstrPrefijo = String.Empty, lstrIdNotaCon_1 = String.Empty
            Dim lstrIdNotaCon_N = String.Empty
            If ldrwIdNotasCon.Length > 0 Then
                lstrPrefijo = ClsPanorama.FobjValorCampo(ldrwIdNotasCon(0)(0), EnuTipoValor.EnuString)
                lstrIdNotaCon_1 = ClsPanorama.FobjValorCampo(ldrwIdNotasCon(0)(1), EnuTipoValor.EnuString)
                lstrIdNotaCon_N = ClsPanorama.FobjValorCampo(ldrwIdNotasCon(ldrwIdNotasCon.Length - 1)(1),
                EnuTipoValor.EnuString)
            End If
            If Not String.IsNullOrEmpty(lstrIdNotaCon_1) Then
                lstrIdUltimasNotasAA = lstrPrefijo & ";" & lstrIdNotaCon_1 & ";" & lstrIdNotaCon_N
            End If
        End If
        GobjPanDat.SControleProcesoObj(False)
        Return lstrIdUltimasNotasAA
    End Function
#End Region
#Region "Anular Facturas automaticas a predios con más de un propietario"
    ''' <summary>
    ''' Devuelve un array list con los numeros de las facturas a los propietarios 
    ''' de un mismo predio, donde se facturo lo mismo y la misma fecha 
    ''' </summary>
    ''' <param name="aobjFactura"></param>
    ''' <returns></returns>
    Friend Function FarlFactsPropietarios(aobjFactura As ClsFactura) As ArrayList
        Dim lstrPref As String, lentIdFact As Integer, lstrNroFac As String
        Dim larlNrosFacs As New ArrayList From {
            aobjFactura.StrNumeroFactura
            }
        If Not String.IsNullOrEmpty(aobjFactura.ObjIdPredioAgrupador_FacStr.ToString) Then
            Dim ldtbIdFacts = FdtbIdFacts(aobjFactura.ObjIdPredioAgrupador_FacStr.ObjValorPro,
                aobjFactura.ObjFechaFacturaDtm.ObjValorPro)
            If ldtbIdFacts.Rows.Count > 1 Then
                For Each ldrwIdFac As DataRow In ldtbIdFacts.Rows
                    lstrPref = ClsPanorama.FobjValorCampo(ldrwIdFac(
                            ClsPrefijo_FactStr.SstrNombreCampoBd), EnuTipoValor.EnuString)
                    lentIdFact = ClsPanorama.FobjValorCampo(ldrwIdFac(
                            ClsIdFacturaEnt.SstrNombreCampoBd), EnuTipoValor.EnuInteger)
                    lstrNroFac = ClsPanorama.FstrNumeroDcto(lstrPref, lentIdFact)
                    If Not larlNrosFacs.Contains(lstrNroFac) Then
                        larlNrosFacs.Add(lstrNroFac)
                    End If
                Next
            End If
        End If
        larlNrosFacs.Sort()
        ' Selecciono las facturas con al menos un servicio facturado igual al servicio
        ' de la factura pasada en el argumento
        Dim lobjFact As New ClsFactura(), lobjValorLlave As Object()
        Dim larlNrosFactRem As New ArrayList
        For Each lstrIdFac As String In larlNrosFacs
            lstrPref = ClsPanorama.FstrPrefijoDcto(lstrIdFac)
            lentIdFact = ClsPanorama.FentIdDcto(lstrIdFac)
            lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentIdFact}
            lobjFact.SAbra(lobjValorLlave)
            If Not FblnSrvFacturadosIguales(aobjFactura, lobjFact) Then
                larlNrosFactRem.Add(lstrIdFac)
            End If
        Next
        For Each lstrIdFacRem As String In larlNrosFactRem
            larlNrosFacs.Remove(lstrIdFacRem)
        Next
        If larlNrosFacs.Count > 1 Then
            larlNrosFactRem.Clear()
            ' Verifico que las facturas sean consecutivas
            Dim i = 0, lentIdFactAct As Integer, lentIdFacAnt As Integer
            For Each lstrNroFacAct As String In larlNrosFacs
                If i = 0 Then
                    lentIdFacAnt = ClsPanorama.FentIdDcto(lstrNroFacAct)
                    i += 1
                Else
                    lentIdFactAct = ClsPanorama.FentIdDcto(lstrNroFacAct)
                    If lentIdFactAct <> lentIdFacAnt + 1 Then
                        larlNrosFactRem.Add(lstrNroFacAct)
                    Else
                        lentIdFacAnt = lentIdFactAct
                    End If
                End If
            Next
            For Each lstrIdFacRem As String In larlNrosFactRem
                larlNrosFacs.Remove(lstrIdFacRem)
            Next
        End If
        Return larlNrosFacs
    End Function
    Private Function FdtbIdFacts(astrIdPredioAgr As String, adtmFecha As Date) As DataTable
        Dim lstrFecha = "'" & ClsPanoramaDat.FstrFechaNormalizada(adtmFecha) & "'"
        Dim lstrTabla = ClsFactura.SstrNombreTabla
        Dim lstrCampSel As String() = {ClsPrefijo_FactStr.SstrNombreCampoBd,
                ClsIdFacturaEnt.SstrNombreCampoBd}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " &
                ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd & " = '" & astrIdPredioAgr & "'" &
                " AND " & ClsFechaFacturaDtm.SstrNombreCampoBd & " = " & lstrFecha
        Dim lstrOrden = {{ClsIdFacturaEnt.SstrNombreCampoBd, "ASC"}}
        Dim ldtbIdFacts = ClsPanorama.FdtbDataTable(lstrTabla, lstrCampSel, lstrOrden, lstrFiltro)
        Return ldtbIdFacts
    End Function
    Private Function FblnSrvFacturadosIguales(aobjFactPrin As ClsFactura,
            aobjFactSec As ClsFactura) As Boolean
        Dim lblnIguales As Boolean, lshrIdAno As Short, lshrIdServicio As Short
        For Each lobjItemFac As ClsItemFactura In aobjFactPrin.ColItemsFactura
            lshrIdAno = lobjItemFac.ObjIdAno_ServicioItemFactShr.ObjValorPro
            lshrIdServicio = lobjItemFac.ObjIdServicio_ItemFactShr.ObjValorPro
            lblnIguales = FblnFacturoServicio(aobjFactSec, lshrIdAno, lshrIdServicio)
            If lblnIguales Then Exit For
        Next
        Return lblnIguales
    End Function
    Private Function FblnFacturoServicio(aobjFact As ClsFactura, ashrIdAno As Short,
            ashrIdServicio As Short) As Boolean
        Dim lblnFacturo As Boolean
        For Each lobjItemFac As ClsItemFactura In aobjFact.ColItemsFactura
            lblnFacturo = lobjItemFac.ObjIdAno_ServicioItemFactShr.ObjValorPro = ashrIdAno AndAlso
                    lobjItemFac.ObjIdServicio_ItemFactShr.ObjValorPro = ashrIdServicio
            If lblnFacturo Then Exit For
        Next
        Return lblnFacturo
    End Function
#End Region
#Region "Factura Manual"
    Friend Sub SGenereFacturaManual(aobjFactura As ClsFactura, ByRef ablnCausoIntMora As Boolean,
            ByRef astrMens As String)
        Dim ldtmfechaFac As Date = aobjFactura.ObjFechaFacturaDtm.ObjValorPro
        Dim lstrIdPreAgr As String = aobjFactura.ObjIdPredioAgrupador_FacStr.ObjValorPro
        Dim ldblIdCliente As Double = aobjFactura.ObjIdCliente_FactDbl.ObjValorPro
        Dim lobjCliente As ClsCliente = aobjFactura.ObjClienteFactura
        Dim lentIdFac As Integer = aobjFactura.ObjIdFacturaEnt.ObjValorPro
        Dim lstrPref As String = aobjFactura.ObjPrefijo_FactStr.ObjValorPro
        Dim lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            GobjPanDat.SInicialiceTransaccion()
            If GobjParametros.ObjFechaUltCausacionGralDtm.ObjValorPro < Date.Today Then
                Dim lstrIdPredAgr As String = aobjFactura.ObjIdPredioAgrupador_FacStr.ToString
                Dim ldtmFecFac As Date = aobjFactura.ObjFechaFacturaDtm.ObjValorPro
                Dim ldecValorInt = lobjCliente.SCauseMora({lstrIdPredAgr}, ldtmFecFac, astrMens)
                ablnCausoIntMora = ldecValorInt > 0
            End If
            Dim lobjEstadoCta As ClsEstadoCuenta
            Dim ldtbItemsFacVivas As DataTable
            If String.IsNullOrEmpty(lstrIdPreAgr) Then
                ldtbItemsFacVivas = FdtbItemsFacVivas(ldblIdCliente, 0)
                lobjEstadoCta = FobjEstCtaClieNoPorSer(lobjCliente, ldtmfechaFac, ldtbItemsFacVivas,
                        lstrPref, lentIdFac)
            Else
                ldtbItemsFacVivas = FdtbItemsFacVivas(ldblIdCliente, lstrIdPreAgr, "A")
                lobjEstadoCta = FobjEstCtaPredNoPorSer(ldblIdCliente, lstrIdPreAgr, ldtmfechaFac,
                        ldtbItemsFacVivas, lstrPref, lentIdFac)
            End If
            aobjFactura.SActualice(True)
            If lobjEstadoCta.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
                lobjEstadoCta.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
            End If
            lobjEstadoCta.ObjPrefijoFac_EstadoStr.ObjValorPro =
                    aobjFactura.ObjPrefijo_FactStr.ObjValorPro
            lobjEstadoCta.ObjIdFactura_EstadoEnt.ObjValorPro =
                    aobjFactura.ObjIdFacturaEnt.ObjValorPro
            lobjEstadoCta.SActualice(True)
            SApliqueAntFacMan(aobjFactura)
            lblnNoHayError = True
        Catch ex As ErrorInesperadoPanDatException
            Throw
        Catch ex As ErrorInesperadoPanLException
            Throw
        Catch ex As ArgumentException
            Throw
        Catch ex As ConexionBdPanException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                GobjPanDat.SConfirmeTransaccion()
                GobjPanDat.SControleProcesoObj(False)
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
#End Region
#Region "Importar Facturas Contingencia"
    Friend Function FblnImportoFrasCon(ByRef astrMens As String) As Boolean
        Dim lblnNoHayError = False, lblnImportoFras = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            Dim lstrArchivo = GstrTrayDatPrg & "PlantillaFrasCon_OrionPLus.xlsx"
            If My.Computer.FileSystem.FileExists(lstrArchivo) Then
                BlnProcesoEspecial = True
                Dim ldtbFras = ClsPanorama.FdtbTablaAccess(lstrArchivo, "", "Plantilla$")
                GobjPanDat.SInicialiceTransaccion()
                lblnImportoFras = FblnImportoFrasCon(ldtbFras, astrMens)
            End If
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
            BlnProcesoEspecial = False
            If lblnNoHayError AndAlso lblnImportoFras Then
                GobjPanDat.SConfirmeTransaccion()
                GobjPanDat.SControleProcesoObj(False)
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
        Return lblnImportoFras
    End Function
    Private Function FblnImportoFrasCon(adtbFras As DataTable, ByRef astrMens As String) As Boolean
        Dim lblnImporto = True
        If FblnEstaArchImpFacOK(adtbFras, astrMens) Then
            SGenereEstadosCtaFactCont(adtbFras)
            Dim lobjCliente = New ClsCliente(EnuModoInstanciaObjDef.EnuUnico)
            Dim ldblIdCliente = 0.0, lstrPreAgr = String.Empty, lstrIdPredio = String.Empty
            Dim lstrPrefFac = String.Empty, lstrPredio = String.Empty, lentIdFac = 0
            Dim ldecValor = 0D, lentIdItemFac = 0, lentIdServicio = 0, lentIdFacNew = 0
            Dim lshrIdAno = 0S
            Dim lentFormaPago = 0, lentMedioPago = 0, lentIdInfCont = 0
            Dim lstrKeyServicio = String.Empty
            Dim ldtmFechaFac = GCDTMFECHANULA, ldtmFechaVen As DateTime = GCDTMFECHANULA
            Dim lstrHora As String = String.Empty
            Dim lstrPref = GobjParametros.FstrPrefijoDoc(EnuIdDocumentoDef.EnuFacturaVenta)
            Dim lobjFactura As New ClsFactura(lstrPref)
            For Each ldrwItemFac As DataRow In adtbFras.Rows
                If Not FblnRegVacio(ldrwItemFac) Then
                    lentIdInfCont = ClsPanorama.FobjValorCampo(ldrwItemFac("NroInformeCont"),
                            EnuTipoValor.EnuInteger)
                    ldblIdCliente = ClsPanorama.FobjValorCampo(ldrwItemFac("IdCliente"),
                            EnuTipoValor.EnuDouble)
                    lstrPreAgr = ClsPanorama.FobjValorCampo(ldrwItemFac("IdPreAgr"),
                            EnuTipoValor.EnuString)
                    lstrPrefFac = ClsPanorama.FobjValorCampo(ldrwItemFac("PrefFactura"),
                            EnuTipoValor.EnuString)
                    lentIdFac = ClsPanorama.FobjValorCampo(ldrwItemFac("IdFactura"),
                            EnuTipoValor.EnuInteger)
                    ldtmFechaFac = ClsPanorama.FobjValorCampo(ldrwItemFac("FechaFac"),
                            EnuTipoValor.EnuDate)
                    ldtmFechaVen = ClsPanorama.FobjValorCampo(ldrwItemFac("FechaVence"),
                            EnuTipoValor.EnuDate)
                    lstrHora = ClsPanorama.FobjValorCampo(ldrwItemFac("HoraFac"),
                            EnuTipoValor.EnuString)
                    lstrIdPredio = ClsPanorama.FobjValorCampo(ldrwItemFac("IdPredio"),
                            EnuTipoValor.EnuString)
                    lshrIdAno = ClsPanorama.FobjValorCampo(ldrwItemFac("IdAnoServicio"),
                            EnuTipoValor.EnuShort)
                    lentIdServicio = ClsPanorama.FobjValorCampo(ldrwItemFac("IdServicio"),
                            EnuTipoValor.EnuInteger)
                    ldecValor = ClsPanorama.FobjValorCampo(ldrwItemFac("Valor"),
                            EnuTipoValor.EnuDecimal)
                    lentFormaPago = ClsPanorama.FobjValorCampo(ldrwItemFac("FormaPago"),
                            EnuTipoValor.EnuInteger)
                    lentMedioPago = ClsPanorama.FobjValorCampo(ldrwItemFac("MedioPago"),
                            EnuTipoValor.EnuInteger)
                    lstrKeyServicio = "0," & lentIdServicio.ToString
                    Dim lentHora As Integer = lstrHora.Split(":")(0)
                    Dim lentMin As Integer = lstrHora.Split(":")(1)
                    If lstrHora.EndsWith("p.m.") Then
                        lentHora += 12
                    End If
                    ldtmFechaFac = ldtmFechaFac.AddHours(lentHora).AddMinutes(lentMin)
                    Dim lobjServicio As ClsServicio = GobjParametros.ColServiciosPer(lstrKeyServicio)
                    If IsNothing(lstrPrefFac) Then lstrPrefFac = String.Empty
                    If IsNothing(lstrPredio) Then lstrPredio = String.Empty
                    If IsNothing(lstrPreAgr) Then lstrPreAgr = String.Empty
                    If lentIdFac <> lentIdFacNew Then
                        If lentIdFacNew > 0 Then
                            lobjFactura.SActualice(True)
                            SEnlaceConEstadoCuenta(lobjFactura)
                        End If
                        lentIdFacNew = lentIdFac
                        lobjFactura.SCreeObj(Nothing)
                        lobjFactura.ObjIdModoFacturacionByt.ObjValorPro = EnuModoFacturacionDef.EnuContingencia
                        lobjFactura.ObjPrefijo_FactStr.ObjValorPro = lstrPrefFac
                        lobjFactura.ObjIdFacturaEnt.ObjValorPro = lentIdFacNew
                        lobjFactura.ObjIdCliente_FactDbl.ObjValorPro = ldblIdCliente
                        lobjFactura.ObjIdPredioAgrupador_FacStr.ObjValorPro = lstrPreAgr
                        lobjFactura.ObjFechaFacturaDtm.ObjValorPro = ldtmFechaFac
                        lobjFactura.ObjFechaVencimientoDtm.ObjValorPro = ldtmFechaVen
                        lobjFactura.ObjFechaGraciaDtm.ObjValorPro = ldtmFechaVen
                        lobjFactura.ObjPieFacturaUno_FactStr.ObjValorPro = GobjParametros.FstrPieFacturaResCon
                        lobjFactura.ObjIdFormaPagoByt.ObjValorPro = lentFormaPago
                        lobjFactura.ObjIdMedioPagoByt.ObjValorPro = lentMedioPago
                        lobjFactura.ObjIdInformeCont_FacEnt.ObjValorPro = lentIdInfCont
                    End If
                    Dim lstrDetalle = lobjServicio.ObjNombreServicioStr.ObjValorPro
                    If lobjServicio.BlnEsCuotaAdministracion Then
                        If lobjServicio.ObjEsAjusteBln.ObjValorPro Then
                            lstrDetalle &= " " & GobjParametros.ObjAnoActual.ObjIdAnoShr.ToString
                        Else
                            lstrDetalle &= " " & GobjParametros.ObjAnoActual.StrNombrePeriodoActual
                        End If
                    End If
                    Dim lobjItemFact As ClsItemFactura = lobjFactura.FobjNuevoItemFactura
                    With lobjItemFact
                        .ObjPeriodo_ItemFactStr.ObjValorPro = GobjParametros.ObjAnoActual.StrIdPeriodoActual
                        .ObjIdAno_ServicioItemFactShr.ObjValorPro = lshrIdAno
                        .ObjIdServicio_ItemFactShr.ObjValorPro = lentIdServicio
                        .ObjFechaGraciaIFDtm.ObjValorPro = ldtmFechaVen
                        .ObjFechaVencimientoIFDtm.ObjValorPro = ldtmFechaVen
                        .ObjIdPredio_ItemFactStr.ObjValorPro = lstrIdPredio
                        .ObjValor_ItemFactDec.ObjValorPro = ldecValor
                        .ObjDetalle_ItemFactStr.ObjValorPro = lstrDetalle
                        .ObjFechaVencimientoIFDtm.ObjValorPro = ldtmFechaVen
                        .ObjFechaGraciaIFDtm.ObjValorPro = ldtmFechaVen
                    End With
                    lobjFactura.SAdicioneNuevoItem(lobjItemFact)
                End If
            Next
            If lobjFactura.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
                lobjFactura.SActualice(True)
                SEnlaceConEstadoCuenta(lobjFactura)
            End If
            GobjPanorama.SRegistreAccionLogApp("clsOrionCop", "Importación Facturas Contingencia")
        Else
            lblnImporto = False
        End If
        Return lblnImporto
    End Function
    Private Shared Function FblnEstaArchImpFacOK(adtbFacturasCon As DataTable,
            ByRef astrMens As String) As Boolean
        Dim lblnEstaOk = (adtbFacturasCon.Rows.Count > 0)
        If Not lblnEstaOk Then
            astrMens = "El Archivo de Excel esta vacio!"
        End If
        If lblnEstaOk Then
            lblnEstaOk = FblnCamposOk(adtbFacturasCon, astrMens)
        End If
        Return lblnEstaOk
    End Function
    Private Shared Function FblnCamposOk(adtbFacturasCon As DataTable,
            ByRef astrMens As String) As Boolean
        Dim lobjCliente As New ClsCliente(EnuModoInstanciaObjDef.EnuUnico)
        Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.EnuUnico)
        Dim lobjFactura = New ClsFactura()
        Dim lobjAno As ClsAno
        Dim lblnEstaOk = False
        Dim lstrNroRes As String, lentIdServicio As Integer, lshrIdAno As Short
        Dim lstrKeyServicio As String, ldecValor As Decimal
        For Each ldrwItemFac As DataRow In adtbFacturasCon.Rows
            If Not FblnRegVacio(ldrwItemFac) Then
                ' Valida número resolución
                lstrNroRes = ClsPanorama.FobjValorCampo(ldrwItemFac("NroResolucion"),
                        EnuTipoValor.EnuString)
                lblnEstaOk = (lstrNroRes = GobjParametros.ObjNumeroResolContiStr.ObjValorPro)
                If Not lblnEstaOk Then
                    astrMens = "El Número de la Resolución no coincide con el registrado en el Sistema!"
                    Exit For
                End If
                ' Valida el cliente y los predios
                lblnEstaOk = FblnClienYPredioOk(ldrwItemFac, lobjCliente, lobjPredio, astrMens)
                If Not lblnEstaOk Then Exit For
                ' Valida Factura y fechas de factura
                lblnEstaOk = FblnNroFacYFechaOk(ldrwItemFac, lobjFactura, astrMens)
                If Not lblnEstaOk Then Exit For
                ' Valida Servicio
                lshrIdAno = ClsPanorama.FobjValorCampo(ldrwItemFac("IdAnoServicio"),
        EnuTipoValor.EnuShort)
                lentIdServicio = ClsPanorama.FobjValorCampo(ldrwItemFac("IdServicio"),
        EnuTipoValor.EnuInteger)
                lstrKeyServicio = lshrIdAno.ToString & "," & lentIdServicio.ToString
                If lshrIdAno = 0 Then
                    lblnEstaOk = GobjParametros.ColServiciosPer.Contains(lstrKeyServicio)
                    If Not lblnEstaOk Then
                        astrMens = "No existe el Servicio Permanente " & lentIdServicio.ToString
                        Exit For
                    End If
                Else
                    lblnEstaOk = (lshrIdAno = GobjParametros.ObjAnoActual.ObjIdAnoShr.ObjValorPro)
                    If Not lblnEstaOk Then
                        astrMens = "El Id. del Año del Servicio tiene Inconsistencias!"
                        Exit For
                    End If
                    lobjAno = GobjParametros.ColAnos(lshrIdAno.ToString)
                    lblnEstaOk = lobjAno.ColServiciosAno.Contains(lstrKeyServicio)
                    If Not lblnEstaOk Then
                        astrMens = "No existe el Servicio del Año " & lentIdServicio.ToString
                        Exit For
                    End If
                End If
                ldecValor = ClsPanorama.FobjValorCampo(ldrwItemFac("Valor"),
        EnuTipoValor.EnuDecimal)
                lblnEstaOk = (ldecValor > 0)
                If Not lblnEstaOk Then
                    Dim lstrPrefFac As String = ClsPanorama.FobjValorCampo(ldrwItemFac("PrefFactura"),
        EnuTipoValor.EnuString)
                    Dim lentIdFac As Integer = ClsPanorama.FobjValorCampo(ldrwItemFac("IdFactura"),
        EnuTipoValor.EnuInteger)
                    Dim lstrNroFac = lstrPrefFac & "-" & lentIdFac.ToString
                    astrMens = "El Valor a facturar debe ser mayor a cero en Factura " & lstrNroFac &
        " e IdServicioPer " & lentIdServicio.ToString
                    Exit For
                End If
            End If
        Next
        Return lblnEstaOk
    End Function
    Private Shared Function FblnClienYPredioOk(adrwItemFact As DataRow, aobjCliente As ClsCliente,
            aobjPredio As ClsPredio, ByRef astrMens As String) As Boolean
        Dim ldblIdCliente As Double = ClsPanorama.FobjValorCampo(adrwItemFact("IdCliente"),
                EnuTipoValor.EnuDouble)
        aobjCliente.SAbra({GshrIdCarpeta, GshrIdCentroUtil, ldblIdCliente})
        Dim lblnEstaOk = aobjCliente.BlnExiste
        If Not lblnEstaOk Then
            astrMens = "El Cliente " & ldblIdCliente.ToString & " no está creado!"
            Return lblnEstaOk
        End If
        Dim lstrPreAgr As String = ClsPanorama.FobjValorCampo(adrwItemFact("IdPreAgr"),
                EnuTipoValor.EnuString)
        Dim lstrPre As String
        If Not String.IsNullOrEmpty(lstrPreAgr) Then
            aobjPredio.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPreAgr})
            lblnEstaOk = aobjPredio.BlnExiste
            If Not lblnEstaOk Then
                astrMens = "El Predio Agrupador " & lstrPreAgr & " no está creado!"
                Return lblnEstaOk
            End If
            lblnEstaOk = (aobjPredio.ObjIdPredioStr.ObjValorPro =
                    aobjPredio.ObjIdPredioAgrupadorStr.ObjValorPro)
            If Not lblnEstaOk Then
                astrMens = "El Predio " & lstrPreAgr & " no es un Predio Agrupador!"
                Return lblnEstaOk
            End If
            lstrPre = ClsPanorama.FobjValorCampo(adrwItemFact("IdPredio"),
                    EnuTipoValor.EnuString)
            lblnEstaOk = Not String.IsNullOrEmpty(lstrPre)
            If Not lblnEstaOk Then
                astrMens = "La Factura con PredioAgrupador " & lstrPreAgr &
                        " debe tener un Predio asociado!"
                Return lblnEstaOk
            End If
            aobjPredio.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPre})
            lblnEstaOk = aobjPredio.BlnExiste
            If Not lblnEstaOk Then
                astrMens = "El Predio " & lstrPreAgr & " no está creado!"
                Return lblnEstaOk
            End If
            lblnEstaOk = (aobjPredio.ObjIdPredioAgrupadorStr.ObjValorPro = lstrPreAgr)
            If Not lblnEstaOk Then
                astrMens = "La PredioAgrupador " & lstrPreAgr & " no agrupa al Predio " &
                        lstrPre & "!"
                Return lblnEstaOk
            End If
        Else
            lstrPre = ClsPanorama.FobjValorCampo(adrwItemFact("IdPredio"), EnuTipoValor.EnuString)
            lblnEstaOk = String.IsNullOrEmpty(lstrPre)
            If Not lblnEstaOk Then
                astrMens = "La Predio " & lstrPre & " no tiene Predio Agrupador!"
                Return lblnEstaOk
            End If
        End If
        Return lblnEstaOk
    End Function
    Private Shared Function FblnNroFacYFechaOk(adrwItemFact As DataRow, aobjFactura As ClsFactura,
            ByRef astrMens As String) As Boolean
        Dim lstrPrefFac As String = ClsPanorama.FobjValorCampo(adrwItemFact("PrefFactura"),
                EnuTipoValor.EnuString)
        Dim lentIdFac As Integer = ClsPanorama.FobjValorCampo(adrwItemFact("IdFactura"),
                EnuTipoValor.EnuInteger)
        Dim lstrNroFac = lstrPrefFac & "-" & lentIdFac.ToString
        If IsNothing(lstrPrefFac) Then lstrPrefFac = String.Empty
        aobjFactura.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefFac, lentIdFac})
        Dim lblnEstaOk = Not aobjFactura.BlnExiste
        If Not lblnEstaOk Then
            astrMens = "La Factura Nro. " & lstrNroFac & " ya existe en el sistema!"
            Return lblnEstaOk
        End If
        lblnEstaOk = (lstrPrefFac = GobjParametros.ObjPrefijoFactContStr.ObjValorPro)
        If Not lblnEstaOk Then
            astrMens = "El Prefijo de la factura " & lstrPrefFac & "-" & lentIdFac.ToString &
    "               no corresponde al prefijo de la Resolución!"
            Return lblnEstaOk
        End If
        lblnEstaOk = (lentIdFac <= GobjParametros.ObjRangoFraConFinEnt.ObjValorPro AndAlso
                lentIdFac >= GobjParametros.ObjRangoFraConIniEnt.ObjValorPro)
        If Not lblnEstaOk Then
            astrMens = "La Factura Nro. " & lstrPrefFac & "-" & lentIdFac.ToString &
        " esta por fuera del Rango autorizado en la Resolución!"
            Return lblnEstaOk
        End If
        Dim ldtmFechaFac As Date = ClsPanorama.FobjValorCampo(adrwItemFact("FechaFac"),
                EnuTipoValor.EnuDate)
        Dim ldtmFechaVence As Date = ClsPanorama.FobjValorCampo(adrwItemFact("FechaVence"),
                EnuTipoValor.EnuDate)
        lblnEstaOk = (ldtmFechaFac >=
                GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo) AndAlso
                (ldtmFechaVence <=
                GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo)
        If Not lblnEstaOk Then
            astrMens = "La Fecha de la Factura " & lstrNroFac & " debe pertenecer al Período Actual!"
            Return lblnEstaOk
        End If
        lblnEstaOk = (ldtmFechaVence >= ldtmFechaFac)
        If Not lblnEstaOk Then
            astrMens = "La Fecha de Vencimiento debe ser posterior a la Fecha de la Factura " &
                    lstrNroFac
            Return lblnEstaOk
        End If
        lblnEstaOk = FblnFechaFacConOk(adrwItemFact, astrMens)
        If Not lblnEstaOk Then
            Return lblnEstaOk
        End If
        Dim ldecValor As Decimal = ClsPanorama.FobjValorCampo(adrwItemFact("Valor"),
                EnuTipoValor.EnuDecimal)
        lblnEstaOk = ldecValor - Int(ldecValor) = 0
        If Not lblnEstaOk Then
            astrMens = "La Valor de Factura " & lstrNroFac & " tiene Centavos"
            Return lblnEstaOk
        End If
        Return lblnEstaOk
    End Function
    Private Shared Function FblnFechaFacConOk(adrwItemFact As DataRow, ByRef astrMens As String) As Boolean
        Dim lblnOk As Boolean
        Dim lentIdInfCon = ClsPanorama.FobjValorCampo(adrwItemFact("NroInformeCont"),
                EnuTipoValor.EnuInteger)
        Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, lentIdInfCon}
        Dim lobjInfCon As New ClsInformeCont(EnuModoInstanciaObjDef.EnuUnico)
        lobjInfCon.SAbra(lobjValorLlave)
        Dim ldtmFecIniCon As DateTime = lobjInfCon.ObjFechaInicioContDtm.ObjValorPro
        Dim ldtmFecFinCon As DateTime = lobjInfCon.ObjFechaFinContDtm.ObjValorPro
        Dim ldtmFecFac As DateTime = ClsPanorama.FobjValorCampo(adrwItemFact("FechaFac"),
EnuTipoValor.EnuDateTime)
        Dim lstrHorFac As String = ClsPanorama.FobjValorCampo(adrwItemFact("HoraFac"),
EnuTipoValor.EnuString)
        Dim lentHora As Integer = CType(lstrHorFac.Split(":")(0), Integer)
        Dim lentMin As Integer = CType(lstrHorFac.Split(":")(1), Integer)
        If lstrHorFac.EndsWith("p.m.") Then
            lentHora += 12
        End If
        ldtmFecFac = ldtmFecFac.AddHours(lentHora).AddMinutes(lentMin)
        lblnOk = (ldtmFecFac >= ldtmFecIniCon AndAlso ldtmFecFac <= ldtmFecFinCon)
        If Not lblnOk Then
            Dim lstrPrefFac As String = ClsPanorama.FobjValorCampo(adrwItemFact("PrefFactura"),
    EnuTipoValor.EnuString)
            Dim lentIdFac As Integer = ClsPanorama.FobjValorCampo(adrwItemFact("IdFactura"),
    EnuTipoValor.EnuInteger)
            Dim lstrNroFac = ClsPanorama.FstrNumeroDcto(lstrPrefFac, lentIdFac)

            astrMens = "La Fecha de la Factura " & lstrNroFac & " está por fuera del Período " &
"de la Contingencia!"
        End If
        Return lblnOk
    End Function
    Private Shared Function FblnRegVacio(ldrw As DataRow) As Boolean
        Dim lentNroInfCon As Integer = ClsPanorama.FobjValorCampo(ldrw("NroInformeCont"),
                EnuTipoValor.EnuInteger)
        Dim lstrNroRes As String = ClsPanorama.FobjValorCampo(ldrw("NroResolucion"),
EnuTipoValor.EnuString)
        Dim lentIdFac As Integer = ClsPanorama.FobjValorCampo(ldrw("IdFactura"),
EnuTipoValor.EnuInteger)
        Dim lblnVacio As Boolean = (lentNroInfCon = 0) AndAlso (lstrNroRes Is Nothing) AndAlso
(lentIdFac = 0)
        Return lblnVacio
    End Function
#End Region
#Region "Funciones y Procedimientos de apoyo"
    Friend Shared Function FobjAnticipo() As ClsAnticipo
        Dim lobjAnticipo As ClsAnticipo
        lobjAnticipo = New ClsAnticipo(EnuModoInstanciaObjDef.EnuNavegable)
        Return lobjAnticipo
    End Function
    Private Shared Function FcolAnticipo(aobjFact As ClsFactura) As Collection
        Dim lcolAnticipos As New Collection
        Dim lobjAnt As ClsAnticipo
        Dim lentIdAnt As Integer
        Dim lstrPredio As String = aobjFact.ObjIdPredioAgrupador_FacStr.ToString
        Dim ldblIdCliente As Double = aobjFact.ObjIdCliente_FactDbl.ObjValorPro
        Dim ldtbAntPorApli = FdtbAnticiposPorAplicar(lstrPredio, ldblIdCliente)
        If ldtbAntPorApli.Rows.Count > 0 Then
            For Each ldrwAnt As DataRow In ldtbAntPorApli.Rows
                lentIdAnt = ldrwAnt(ClsIdAnticipoEnt.SstrNombreCampoBd)
                lobjAnt = New ClsAnticipo(EnuModoInstanciaObjDef.EnuUnico)
                lobjAnt.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lentIdAnt})
                lcolAnticipos.Add(lobjAnt)
            Next
        End If
        Return lcolAnticipos
    End Function
    Friend Shared Function FstrNombreModoFacturacion(aenuModoFacturacion As EnuModoFacturacionDef)
        Dim lstrNombre = String.Empty
        Select Case aenuModoFacturacion
            Case EnuModoFacturacionDef.None
                lstrNombre = "Ninguno"
            Case EnuModoFacturacionDef.EnuManual
                lstrNombre = My.Resources.MFManual
            Case EnuModoFacturacionDef.EnuSistema
                lstrNombre = My.Resources.MFSistema
            Case EnuModoFacturacionDef.EnuImportada
                lstrNombre = My.Resources.MFImportada
            Case EnuModoFacturacionDef.EnuContingencia
                lstrNombre = My.Resources.MFContingencia
        End Select
        Return lstrNombre
    End Function
    ''' <summary>
    ''' Devuelve una tabla con datos de los items del programa de facturación que tienen saldo mayor a cero,
    ''' pertenecena la ubicación actual y el año del servicio es cero o, menor o igual al año actual.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>    
    ''' <summary>
    ''' devuelve un string que contiene el prefijo, el numero de la primera y 
    ''' el número de la ultima factura generada or el programa (automática)
    ''' separadas por un punto y coma.
    ''' </summary>
    Friend Shared Function FstrIdUltimasFras(ablnContingencia As Boolean) As String
        GobjPanDat.SControleProcesoObj(True)
        Dim lstrUltimasFacturas = String.Empty
        Dim lstbSql As New StringBuilder
        Dim ldtmFecha = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
        Dim lstrFecha = ClsPanoramaDat.FstrFechaNormalizada(ldtmFecha)
        lstrFecha = " '" & lstrFecha & "'"
        If ldtmFecha <> GCDTMFECHANULA Then
            ' Campos a seleccionar
            lstbSql.Clear()
            With lstbSql
                .Append(ClsPrefijo_FactStr.SstrNombreCampoBd).Append(CMSTRCOMA)
                .Append(ClsIdFacturaEnt.SstrNombreCampoBd)
            End With
            Dim lstrCamposSelect = {lstbSql.ToString}
            ' Filtro
            With lstbSql
                .Clear()
                .Append(StrFiltroUbicacion).Append(" AND ")
                .Append(ClsFechaFacturaDtm.SstrNombreCampoBd).Append(" >= ")
                .Append(lstrFecha).Append(" AND ")
                .Append(ClsIdModoFacturacionByt.SstrNombreCampoBd).Append(" = ")
                If ablnContingencia Then
                    .Append(EnuModoFacturacionDef.EnuContingencia)
                Else
                    .Append(EnuModoFacturacionDef.EnuSistema)
                End If
                .Append(" AND ").Append(ClsPrefijo_FactStr.SstrNombreCampoBd).Append(" <> '")
                .Append(GCSTRPREFPREFACTURA).Append("'")
            End With
            Dim lstrFiltro = lstbSql.ToString
            Dim ldtbInfFras = ClsPanorama.FdtbDataTable(ClsFactura.SstrNombreTabla,
                    lstrCamposSelect, {{ClsIdFacturaEnt.SstrNombreCampoBd, "ASC"}},
                    lstrFiltro)
            Dim ldrwFacturas = ldtbInfFras.Select()
            Dim lstrPref = String.Empty, lentIdFraIni = 0, lentIdFraFin = 0
            If ldrwFacturas.Length > 0 Then
                Dim ldrwFra As DataRow = ldrwFacturas(0)
                lstrPref = ClsPanorama.FobjValorCampo(ldrwFra(0), EnuTipoValor.EnuString)
                lentIdFraIni = ClsPanorama.FobjValorCampo(ldrwFra(1),
                        EnuTipoValor.EnuInteger)
                ldrwFra = ldrwFacturas(ldrwFacturas.Length - 1)
                lentIdFraFin = ClsPanorama.FobjValorCampo(ldrwFra(1),
                         EnuTipoValor.EnuInteger)
                With lstbSql
                    .Clear().Append(lstrPref).Append(";")
                    .Append(lentIdFraIni.ToString).Append(";").Append(lentIdFraFin.ToString)
                End With
                lstrUltimasFacturas = lstbSql.ToString
            End If
        End If
        GobjPanDat.SControleProcesoObj(False)
        Return lstrUltimasFacturas
    End Function
    ''' <summary>
    ''' Indica la cantidad de numeros disponibles para facturas según la Resolución de La Dian
    ''' </summary>
    ''' <returns></returns>
    Friend Shared Function FentCantNumerosFacDisponibles() As Integer
        Dim lentCan As Integer, lstrTabla = ClsFactura.SstrNombreTabla
        Dim lstrPrefijoFac As String = GobjParametros.FstrPrefijoDoc(
                EnuTipoDocOri.EnuFactura)
        Dim lstrFiltroFac = StrFiltroUbicacion & " AND " &
                ClsPrefijo_FactStr.SstrNombreCampoBd & " = '" & lstrPrefijoFac & "'"
        Dim lentIdUltFac As Integer = ClsPanorama.FobjUltimaIdNumericaObjeto(lstrTabla,
        ClsIdFacturaEnt.SstrNombreCampoBd, EnuTipoValor.EnuInteger,
        lstrFiltroFac)
        Dim lentLimiteSupRes As Integer = GobjParametros.ObjRangoFraFinEnt.ObjValorPro
        lentCan = lentLimiteSupRes - lentIdUltFac
        Return lentCan
    End Function
    ''' <summary>
    ''' Indica la cantidad de días que restan para el vencimiento de la Reolución de la Dian
    ''' </summary>
    ''' <returns></returns>
    Friend Shared Function FentCantDiasVigenciaRes() As Integer
        Dim ldtmFechaVenRes As Date = GobjParametros.ObjFechaVenceResolFactDtm.ObjValorPro
        Dim lentDias As Integer = ClsPanorama.FentDiasEntreFechas(Date.Today, ldtmFechaVenRes)
        Return lentDias
    End Function
    Friend Shared Function FblnHayResVigente() As Boolean
        Dim ldtmFechaVenRes As Date = GobjParametros.ObjFechaVenceResolFactDtm.ObjValorPro
        Dim ldtmFechaRes As Date = GobjParametros.ObjFechaResolucionFactDtm.ObjValorPro
        Dim lblnHayRes As Boolean = Date.Now >= ldtmFechaRes AndAlso ldtmFechaVenRes >= Date.Now
        Return lblnHayRes
    End Function
    ''' <summary>
    ''' devuelve un string que contiene los numeros de la primera y la ultima cta cobro generadas, 
    ''' separadas por un punto y coma.
    ''' </summary>
    Friend Shared Function FstrIdUltimasCtasCobro() As String
        GobjPanDat.SControleProcesoObj(True)
        Dim lstrIdUltimasCtasCobro = String.Empty
        Dim lstbSql As New StringBuilder
        Dim ldtmFecha = GobjParametros.ObjAnoActual.ObjFechaUltEstadoCtaDtm.ObjValorPro
        Dim lstrFecha = ClsPanoramaDat.FstrFechaNormalizada(ldtmFecha)
        Dim ldrwIdCtasCobro() As DataRow = Nothing
        If ldtmFecha <> GCDTMFECHANULA Then
            ' Campos a seleccionar
            Dim lstrCamposSelect = {ClsIdEstadoCuentaEnt.SstrNombreCampoBd}
            ' Filtro
            With lstbSql
                .Clear()
                .Append(StrFiltroUbicacion).Append(" AND ")
                .Append(ClsFechaEstadoDtm.SstrNombreCampoBd).Append(" = '").Append(lstrFecha)
                .Append("' AND ").Append(ClsIdFactura_EstadoEnt.SstrNombreCampoBd).Append(" = 0")
            End With
            Dim lstrFiltro = lstbSql.ToString
            Dim ldtbIdCtasCobro = ClsPanorama.FdtbDataTable(ClsEstadoCuenta.SstrNombreTabla,
                    lstrCamposSelect, {{ClsIdEstadoCuentaEnt.SstrNombreCampoBd, "ASC"}}, lstrFiltro)
            ldrwIdCtasCobro = ldtbIdCtasCobro.Select()
            Dim lstrIdCtaCobro_1 = String.Empty, lstrIdCtaCobro_N = String.Empty
            If ldrwIdCtasCobro.Length > 0 Then
                lstrIdCtaCobro_1 = CType(ldrwIdCtasCobro(0)(0), String)
                lstrIdCtaCobro_N = CType(ldrwIdCtasCobro(ldrwIdCtasCobro.Length - 1)(0), String)
            End If
            If Not String.IsNullOrEmpty(lstrIdCtaCobro_1) Then
                lstrIdUltimasCtasCobro = lstrIdCtaCobro_1 & ";" & lstrIdCtaCobro_N
            End If
        End If
        GobjPanDat.SControleProcesoObj(False)
        Return lstrIdUltimasCtasCobro
    End Function
    ''' <summary>
    ''' devuelve un string que contiene elprefijo y los numeros de la primera y las ultimas notas db
    ''' generadas en el período actual, separadas por un punto y coma.
    ''' </summary>
    Friend Shared Function FstrIdUltimasNotasDb() As String
        GobjPanDat.SControleProcesoObj(True)
        Dim lstrIdUltimasNotasDb = String.Empty
        Dim lstbSql As New StringBuilder
        Dim ldtmFecha = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
        Dim lstrFecha = "'" & ClsPanoramaDat.FstrFechaNormalizada(ldtmFecha) & "'"
        Dim ldrwIdNotasDb() As DataRow = Nothing
        If ldtmFecha <> GCDTMFECHANULA Then
            ' Campos a seleccionar
            Dim lstrCamposSelect = {ClsPrefijo_NotaDbStr.SstrNombreCampoBd,
                    ClsIdNotaDbEnt.SstrNombreCampoBd}
            ' Filtro
            With lstbSql
                .Clear()
                .Append(StrFiltroUbicacion).Append(" AND ")
                .Append(ClsFecha_NotaDbDtm.SstrNombreCampoBd).Append(" >= ").Append(lstrFecha)
            End With
            Dim lstrFiltro = lstbSql.ToString
            Dim ldtbIdNotasDb = ClsPanorama.FdtbDataTable(ClsNotaDb.SstrNombreTabla, lstrCamposSelect,
                    {{ClsIdNotaDbEnt.SstrNombreCampoBd, "ASC"}}, lstrFiltro)
            ldrwIdNotasDb = ldtbIdNotasDb.Select()
            Dim lstrPrefijo = String.Empty, lstrIdNotaDb_1 = String.Empty
            Dim lstrIdNotaDb_N = String.Empty
            If ldrwIdNotasDb.Length > 0 Then
                lstrPrefijo = ClsPanorama.FobjValorCampo(ldrwIdNotasDb(0)(0),
                        EnuTipoValor.EnuString)
                lstrIdNotaDb_1 = ClsPanorama.FobjValorCampo(ldrwIdNotasDb(0)(1),
                        EnuTipoValor.EnuString)
                lstrIdNotaDb_N = ClsPanorama.FobjValorCampo(ldrwIdNotasDb(ldrwIdNotasDb.Length - 1)(1),
                        EnuTipoValor.EnuString)
            End If
            If Not String.IsNullOrEmpty(lstrIdNotaDb_1) Then
                lstrIdUltimasNotasDb = lstrPrefijo & ";" & lstrIdNotaDb_1 & ";" & lstrIdNotaDb_N
            End If
        End If
        GobjPanDat.SControleProcesoObj(False)
        Return lstrIdUltimasNotasDb
    End Function
    Friend Shared Function FdtbPrefacturas() As DataTable
        Dim lstrFiltro = StrFiltroUbicacion & " AND " & ClsEsPreFacturaBln.SstrNombreCampoBd &
                " = True"
        Dim lstrIndice As String(,) = {{ClsPrefijo_FactStr.SstrNombreCampoBd, "ASC"},
                {ClsIdFacturaEnt.SstrNombreCampoBd, "ASC"}}
        Dim ldtbPrefacturas As DataTable = ClsPanorama.FdtbDataTable(ClsFactura.SstrNombreTabla,
                {"*"}, lstrIndice, lstrFiltro)
        Return ldtbPrefacturas
    End Function
    ''' <summary>
    ''' Devuelve un DataTable que contiene las cuentas bancarias y la cuenta de Caja.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Shared Function FdtbCuentasIngresos() As DataTable
        Dim lstrTabla = ClsCuentaBanco.SstrNombreTabla
        Dim lstrCamposSelect = {"*"}
        Dim lstrIndice = {{ClsIdCuentaBancoShr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " & ClsEstaActivaBln.SstrNombreCampoBd & " = True"
        Dim ldtbCuentasIngresos = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamposSelect, lstrIndice,
                lstrFiltro)
        Dim ldrwNinguno As DataRow = ldtbCuentasIngresos.NewRow
        ldrwNinguno(StrCampoCarpeta) = 1
        ldrwNinguno(StrCampoCentroUtil) = 1
        ldrwNinguno(ClsIdCuentaBancoShr.SstrNombreCampoBd) = -2
        ldrwNinguno(ClsIdCtaContabilidadStr.SstrNombreCampoBd) = String.Empty
        ldrwNinguno(ClsNombreBancoStr.SstrNombreCampoBd) = "<Ninguno>"
        ldrwNinguno(ClsNumeroCuentaStr.SstrNombreCampoBd) = String.Empty
        ldtbCuentasIngresos.Rows.InsertAt(ldrwNinguno, 0)
        '
        Dim ldrwEfectivo As DataRow = ldtbCuentasIngresos.NewRow
        ldrwEfectivo(StrCampoCarpeta) = 1
        ldrwEfectivo(StrCampoCentroUtil) = 1
        ldrwEfectivo(ClsIdCuentaBancoShr.SstrNombreCampoBd) = 0
        ldrwEfectivo(ClsIdCtaContabilidadStr.SstrNombreCampoBd) =
                GobjParametros.ObjIdCtaCajaStr.ObjValorPro
        ldrwEfectivo(ClsNombreBancoStr.SstrNombreCampoBd) = "CAJA"
        ldrwEfectivo(ClsNumeroCuentaStr.SstrNombreCampoBd) = String.Empty
        ldtbCuentasIngresos.Rows.InsertAt(ldrwEfectivo, 1)
        '
        If Not String.IsNullOrEmpty(GobjParametros.ObjIdCtaIngPorIdentificarStr.ToString) Then
            Dim ldrwIngSinIdentificat As DataRow = ldtbCuentasIngresos.NewRow
            ldrwIngSinIdentificat(StrCampoCarpeta) = 1
            ldrwIngSinIdentificat(StrCampoCentroUtil) = 1
            ldrwIngSinIdentificat(ClsIdCuentaBancoShr.SstrNombreCampoBd) = -1
            ldrwIngSinIdentificat(ClsIdCtaContabilidadStr.SstrNombreCampoBd) =
                    GobjParametros.ObjIdCtaIngPorIdentificarStr.ObjValorPro
            ldrwIngSinIdentificat(ClsNombreBancoStr.SstrNombreCampoBd) = "INGRESOS POR IDENTIFICAR"
            ldrwIngSinIdentificat(ClsNumeroCuentaStr.SstrNombreCampoBd) = String.Empty
            ldtbCuentasIngresos.Rows.InsertAt(ldrwIngSinIdentificat, ldtbCuentasIngresos.Rows.Count)
        End If
        Return ldtbCuentasIngresos
    End Function
    ''' <summary>
    ''' Indica si la Copropiedad actual ya tiene creado el registro en OriCentrosUtilidad
    ''' </summary>
    ''' <returns></returns>
    Friend Shared Function FblnExisteRegCenUtil() As Boolean
        Dim lstrFiltro = StrFiltroUbicacion
        Dim ldtbCenutil = ClsPanorama.FdtbDataTable(ClsCentroUtilOriCop.SstrNombreTabla, {"*"},
                {{}}, lstrFiltro)
        Return (ldtbCenutil.Rows.Count > 0)
    End Function
#End Region
#Region "Procedimientos eFac"
    Friend Shared Function FdtbFacsPorProcesar() As DataTable
        Dim lstrTabla = ClsFactura.SstrNombreTabla
        Dim lstrCampSele = {EnuTipoDocOri.EnuFactura & " AS TipoDocu",
                ClsPrefijo_FactStr.SstrNombreCampoBd & " AS Prefijo",
                ClsIdFacturaEnt.SstrNombreCampoBd & " AS IdDocu",
                ClsIdEstadoEDocEnt.SstrNombreCampoBd}
        Dim lstrOrden = {{ClsIdEstadoEDocEnt.SstrNombreCampoBd, "ASC"}, {"Prefijo", "ASC"},
                {"IdDocu", "ASC"}}
        Dim lstrFiltro = StrFiltroUbicacion & " AND (" &
                ClsIdEstadoEDocEnt.SstrNombreCampoBd &
                " = " & EnuEstadoEDoc.EnuNoReg & " OR " &
                ClsIdEstadoEDocEnt.SstrNombreCampoBd & " = " & EnuEstadoEDoc.EnuEnProceso &
                " OR " & ClsIdEstadoEDocEnt.SstrNombreCampoBd & " = " &
                EnuEstadoEDoc.EnuRegi & ")"
        Dim ldtbResu = ClsPanorama.FdtbDataTable(lstrTabla, lstrCampSele, lstrOrden, lstrFiltro)
        Return ldtbResu
    End Function
    Friend Shared Function FdtbNotasPorProcesar() As DataTable
        Dim lstrTabla = ClsNotaDb.SstrNombreTabla
        Dim lstrCampSele = {EnuTipoDocOri.EnuNotaDb & " AS TipoDocu",
                ClsPrefijo_NotaDbStr.SstrNombreCampoBd & " AS Prefijo",
                ClsIdNotaDbEnt.SstrNombreCampoBd & " AS IdDocu",
                ClsIdEstadoEDocEnt.SstrNombreCampoBd}
        Dim lstrOrden = {{ClsIdEstadoEDocEnt.SstrNombreCampoBd, "ASC"},
                {ClsIdNotaDbEnt.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = StrFiltroUbicacion & " AND (" & ClsIdEstadoEDocEnt.SstrNombreCampoBd &
                " = " & EnuEstadoEDoc.EnuNoReg & " OR " &
                ClsIdEstadoEDocEnt.SstrNombreCampoBd & " = " & EnuEstadoEDoc.EnuEnProceso &
                " OR " & ClsIdEstadoEDocEnt.SstrNombreCampoBd & " = " &
                EnuEstadoEDoc.EnuRegi & ")"
        Dim lstrSqlNdb = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTabla, lstrCampSele,
                lstrOrden, lstrFiltro, {})
        '
        lstrTabla = ClsNotaCr.SstrNombreTabla
        lstrCampSele = {EnuTipoDocOri.EnuNotaCr & " AS TipoDocu",
                ClsPrefijo_NotaCrStr.SstrNombreCampoBd & " AS Prefijo",
                ClsIdNotaCrEnt.SstrNombreCampoBd & " AS IdDocu",
                ClsIdEstadoEDocEnt.SstrNombreCampoBd}
        lstrOrden = {{ClsIdEstadoEDocEnt.SstrNombreCampoBd, "ASC"},
                {ClsIdNotaCrEnt.SstrNombreCampoBd, "ASC"}}
        Dim lstrSqlNCr = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTabla, lstrCampSele,
                lstrOrden, lstrFiltro, {})
        '
        lstrTabla = ClsNotaReversionCr.SstrNombreTabla
        lstrCampSele = {EnuTipoDocOri.EnuNotaRevCr & " AS TipoDocu",
                ClsPrefijo_NotaReversaCrStr.SstrNombreCampoBd & " AS Prefijo",
                ClsIdNotaReversaCrEnt.SstrNombreCampoBd & " AS IdDocu",
                ClsIdEstadoEDocEnt.SstrNombreCampoBd}
        lstrOrden = {{ClsIdEstadoEDocEnt.SstrNombreCampoBd, "ASC"},
                {ClsIdNotaReversaCrEnt.SstrNombreCampoBd, "ASC"}}
        Dim lstrSqlNRCr = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTabla, lstrCampSele,
                lstrOrden, lstrFiltro, {})
        '
        lstrTabla = ClsNotaCon.SstrNombreTabla
        lstrCampSele = {EnuTipoDocOri.EnuNotaCon & " AS TipoDocu",
                ClsPrefijo_NotaConStr.SstrNombreCampoBd & " AS Prefijo",
                ClsIdNotaConEnt.SstrNombreCampoBd & " AS IdDocu",
                ClsIdEstadoEDocEnt.SstrNombreCampoBd}
        lstrOrden = {{ClsIdEstadoEDocEnt.SstrNombreCampoBd, "ASC"},
                {ClsIdNotaConEnt.SstrNombreCampoBd, "ASC"}}
        Dim lstrSqlNCon = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTabla, lstrCampSele,
                lstrOrden, lstrFiltro, {})
        Dim lstrSqlRes = "(" & lstrSqlNdb & ") UNION ALL (" & lstrSqlNCr & ") UNION ALL (" &
                lstrSqlNRCr & ") UNION ALL (" & lstrSqlNCon & ")"
        Dim ldtbNotas = ClsPanorama.FdtbDataTable(lstrSqlRes)
        Return ldtbNotas
    End Function
    Friend Shared Function FdtbDocsEstadoCero() As DataTable
        ' ExpSqlFacturas
        Dim lstrCamposSel As String
        Dim ldtmFechaHasta As Date = Today.AddDays(-30)
        Dim lstrFechaDesde = ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaHasta.ToString)
        Dim lstrFilGen = StrFiltroUbicacion & " AND " & ClsIdEstadoEDocEnt.SstrNombreCampoBd &
                " = " & EnuEstadoEDoc.EnuErrorFtp
        Dim lstrTabla = ClsFactura.SstrNombreTabla
        Dim lstrCampSel = {ClsPrefijo_FactStr.SstrNombreCampoBd & " AS Pref",
                ClsIdFacturaEnt.SstrNombreCampoBd & " AS IdDoc",
                ClsFechaFacturaDtm.SstrNombreCampoBd & " AS Fecha",
                EnuTipoDocOri.EnuFactura & " AS TipoDoc"}
        Dim lstrFilDoc = lstrFilGen & " AND " & ClsFechaFacturaDtm.SstrNombreCampoBd & " >= " &
                lstrFechaDesde
        Dim lstrOrden As String = " ORDER BY TipoDoc ASC, Fecha ASC"
        Dim lstrExpFacs = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTabla, lstrCampSel, {{"", ""}},
                lstrFilDoc, {})

        lstrTabla = ClsNotaDb.SstrNombreTabla
        lstrCampSel = {ClsPrefijo_NotaDbStr.SstrNombreCampoBd & " AS Pref",
                ClsIdNotaDbEnt.SstrNombreCampoBd & " AS IdDoc",
                ClsFecha_NotaDbDtm.SstrNombreCampoBd & " AS Fecha",
                EnuTipoDocOri.EnuNotaDb & " AS TipoDoc"}
        lstrFilDoc = lstrFilGen & " AND " & ClsFecha_NotaDbDtm.SstrNombreCampoBd & " >= " &
                lstrFechaDesde
        Dim lstrExpNsDb = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTabla, lstrCampSel, {{"", ""}},
                lstrFilDoc, {})

        lstrTabla = ClsNotaCr.SstrNombreTabla
        lstrCampSel = {ClsPrefijo_NotaCrStr.SstrNombreCampoBd & " AS Pref",
                ClsIdNotaCrEnt.SstrNombreCampoBd & " AS IdDoc",
                ClsFecha_NotaCrDtm.SstrNombreCampoBd & " AS Fecha",
                EnuTipoDocOri.EnuNotaCr & " AS TipoDoc"}
        lstrFilDoc = lstrFilGen & " AND " & ClsFecha_NotaCrDtm.SstrNombreCampoBd & " >= " &
                lstrFechaDesde
        Dim lstrExpNsCr = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTabla, lstrCampSel, {{"", ""}},
                lstrFilDoc, {})

        lstrTabla = ClsNotaReversionCr.SstrNombreTabla
        lstrCamposSel = "IF(" & ClsPrefijo_NotaReversaCrStr.SstrNombreCampoBd &
                " = '', 'RCR', " & ClsPrefijo_NotaReversaCrStr.SstrNombreCampoBd & ") AS Pref, " &
                ClsIdNotaReversaCrEnt.SstrNombreCampoBd & " AS IdDoc, " &
                ClsFecha_NotaReversaCrDtm.SstrNombreCampoBd & " AS Fecha, " &
                EnuTipoDocOri.EnuNotaRevCr & " AS TipoDoc"
        lstrFilDoc = lstrFilGen & " AND " & ClsFecha_NotaReversaCrDtm.SstrNombreCampoBd & " >= " &
                lstrFechaDesde
        Dim lstrExpNsRCr = "SELECT " & lstrCamposSel & " FROM " & lstrTabla & " WHERE " & lstrFilDoc

        lstrTabla = ClsNotaCon.SstrNombreTabla
        lstrCamposSel = "IF(" & ClsPrefijo_NotaConStr.SstrNombreCampoBd &
                        " = '', 'NCON', " & ClsPrefijo_NotaConStr.SstrNombreCampoBd & ") AS Pref, " &
                ClsIdNotaConEnt.SstrNombreCampoBd & " AS IdDoc, " &
                ClsFecha_NotaConDtm.SstrNombreCampoBd & " AS Fecha, " &
                EnuTipoDocOri.EnuNotaCon & " AS TipoDoc"
        lstrFilDoc = lstrFilGen & " AND " & ClsFecha_NotaConDtm.SstrNombreCampoBd & " >= " &
                lstrFechaDesde
        Dim lstrExpNsCon = "SELECT " & lstrCamposSel & " FROM " & lstrTabla & " WHERE " & lstrFilDoc
        Dim lstrSql As String = "(" & lstrExpFacs & ") UNION ALL (" &
                lstrExpNsDb & ") UNION ALL (" & lstrExpNsCr & ") UNION ALL (" &
                lstrExpNsRCr & ") UNION ALL (" & lstrExpNsCon & ")" & lstrOrden
        Dim ldtbResu = ClsPanorama.FdtbDataTable(lstrSql)
        Return ldtbResu
    End Function
    ''' <summary>
    ''' Devuelve un datatable con la identificacion de los documentos electrónicos
    ''' según el tipo de documento y el estado del documento pasados en los aergumentos
    ''' </summary>
    ''' <param name="aenuTipoDoc">Tipo de Documento a tener en cuenta</param>
    ''' <param name="aenuEstadoEDoc">Estado del Documento </param>
    ''' <returns></returns>
    Private Shared Function FdtbDocsEFac(aenuTipoDoc As EnuTipoDocOri,
            aenuEstadoEDoc As EnuEstadoEDoc) As DataTable
        Dim lstrTabla As String, lstrCamSel As String, ldtbResu As DataTable = Nothing
        Dim lstrCamposSelect As String(), lstrExpSql As String
        Dim lstrOrden As String(,)
        Dim lstrFiltro = StrFiltroUbicacion & " AND " & ClsIdEstadoEDocEnt.SstrNombreCampoBd &
                " = " & aenuEstadoEDoc
        Select Case aenuTipoDoc
            Case EnuTipoDocOri.EnuFactura
                lstrTabla = ClsFactura.SstrNombreTabla
                lstrCamposSelect = {ClsPrefijo_FactStr.SstrNombreCampoBd,
                        ClsIdFacturaEnt.SstrNombreCampoBd, ClsFechaFacturaDtm.SstrNombreCampoBd &
                        " AS Fecha"}
                lstrOrden = {{"Fecha", "ASC"}}
                ldtbResu = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamposSelect, lstrOrden,
                        lstrFiltro)
            Case EnuTipoDocOri.EnuNotaDb
                lstrTabla = ClsNotaDb.SstrNombreTabla
                lstrCamposSelect = {ClsPrefijo_NotaDbStr.SstrNombreCampoBd,
                        ClsIdNotaDbEnt.SstrNombreCampoBd, ClsFecha_NotaDbDtm.SstrNombreCampoBd &
                        " AS Fecha"}
                lstrOrden = {{ClsFecha_NotaDbDtm.SstrNombreCampoBd, "ASC"}}
                ldtbResu = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamposSelect, lstrOrden,
                        lstrFiltro)
            Case EnuTipoDocOri.EnuNotaCr
                lstrTabla = ClsNotaCr.SstrNombreTabla
                lstrCamposSelect = {ClsPrefijo_NotaCrStr.SstrNombreCampoBd,
                        ClsIdNotaCrEnt.SstrNombreCampoBd, ClsFecha_NotaCrDtm.SstrNombreCampoBd}
                lstrOrden = {{ClsFecha_NotaCrDtm.SstrNombreCampoBd, "ASC"}}
                ldtbResu = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamposSelect, lstrOrden,
                        lstrFiltro)
            Case EnuTipoDocOri.EnuNotaRevCr
                lstrTabla = ClsNotaReversionCr.SstrNombreTabla
                lstrCamSel = "IF(" & ClsPrefijo_NotaReversaCrStr.SstrNombreCampoBd &
                        " = '', 'RCR', " & ClsPrefijo_NotaReversaCrStr.SstrNombreCampoBd & "), " &
                        ClsIdNotaReversaCrEnt.SstrNombreCampoBd & ", " &
                        ClsFecha_NotaReversaCrDtm.SstrNombreCampoBd
                lstrExpSql = "SELECT " & lstrCamSel & " FROM " & lstrTabla & " WHERE " & lstrFiltro &
                        " ORDER BY " & ClsFecha_NotaReversaCrDtm.SstrNombreCampoBd & " ASC "
                ldtbResu = ClsPanorama.FdtbDataTable(lstrExpSql)
            Case EnuTipoDocOri.EnuNotaCon
                lstrTabla = ClsNotaCon.SstrNombreTabla
                lstrCamSel = "IF(" & ClsPrefijo_NotaConStr.SstrNombreCampoBd &
                        " = '', 'NCON', " & ClsPrefijo_NotaConStr.SstrNombreCampoBd & "), " &
                        ClsIdNotaConEnt.SstrNombreCampoBd & ", " &
                        ClsFecha_NotaConDtm.SstrNombreCampoBd & " AS Fecha"
                lstrExpSql = "SELECT " & lstrCamSel & " FROM " & lstrTabla & " WHERE " & lstrFiltro &
                        " ORDER BY " & ClsFecha_NotaConDtm.SstrNombreCampoBd & " ASC "
                ldtbResu = ClsPanorama.FdtbDataTable(lstrExpSql)
        End Select
        Return ldtbResu
    End Function
    Friend Shared Sub SHabiliteFras()
        Dim ldtbFactInvalidas = FdtbDocsEFac(EnuTipoDocOri.EnuFactura, EnuEstadoEDoc.EnuInvalida)
        Dim lstrPref As String, lentIdFact As Integer
        Dim lobjFac As New ClsFactura()
        For Each ldrwIdFac As DataRow In ldtbFactInvalidas.Rows
            lstrPref = ClsPanorama.FobjValorCampo(ldrwIdFac(ClsPrefijo_FactStr.SstrNombreCampoBd),
                EnuTipoValor.EnuString)
            lentIdFact = ClsPanorama.FobjValorCampo(ldrwIdFac(ClsIdFacturaEnt.SstrNombreCampoBd),
                EnuTipoValor.EnuInteger)
            lobjFac.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentIdFact})
            lobjFac.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
            lobjFac.ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuEnProceso
            lobjFac.SActualice(True)
        Next
    End Sub
    Friend Shared Sub SHabiliteNotas()
        SHabiliteNotas(EnuTipoDocOri.EnuNotaDb)
        SHabiliteNotas(EnuTipoDocOri.EnuNotaCr)
        SHabiliteNotas(EnuTipoDocOri.EnuNotaRevCr)
        SHabiliteNotas(EnuTipoDocOri.EnuNotaCon)
    End Sub
    Private Shared Sub SHabiliteNotas(aenuTipoDoc As EnuTipoDocOri)
        Dim ldtbNotasInvalidas = FdtbDocsEFac(aenuTipoDoc, EnuEstadoEDoc.EnuInvalida)
        Dim lstrPref As String, lentIdNota As Integer
        For Each ldrwIdFac As DataRow In ldtbNotasInvalidas.Rows
            Select Case aenuTipoDoc
                Case EnuTipoDocOri.EnuNotaCon
                    lstrPref = ClsPanorama.FobjValorCampo(ldrwIdFac(ClsPrefijo_NotaConStr.SstrNombreCampoBd),
                            EnuTipoValor.EnuString)
                    lentIdNota = ClsPanorama.FobjValorCampo(ldrwIdFac(ClsIdNotaConEnt.SstrNombreCampoBd),
                            EnuTipoValor.EnuInteger)
                    Dim lobjNota As New ClsNotaCon()
                    lobjNota.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentIdNota})
                    lobjNota.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
                    lobjNota.ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuEnProceso
                    lobjNota.SActualice(True)
                Case EnuTipoDocOri.EnuNotaCr
                    lstrPref = ClsPanorama.FobjValorCampo(ldrwIdFac(ClsPrefijo_NotaCrStr.SstrNombreCampoBd),
                            EnuTipoValor.EnuString)
                    lentIdNota = ClsPanorama.FobjValorCampo(ldrwIdFac(ClsIdNotaCrEnt.SstrNombreCampoBd),
                            EnuTipoValor.EnuInteger)
                    Dim lobjNota As New ClsNotaCr()
                    lobjNota.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentIdNota})
                    lobjNota.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
                    lobjNota.ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuEnProceso
                    lobjNota.SActualice(True)
                Case EnuTipoDocOri.EnuNotaDb
                    lstrPref = ClsPanorama.FobjValorCampo(ldrwIdFac(ClsPrefijo_NotaDbStr.SstrNombreCampoBd),
                            EnuTipoValor.EnuString)
                    lentIdNota = ClsPanorama.FobjValorCampo(ldrwIdFac(ClsIdNotaDbEnt.SstrNombreCampoBd),
                            EnuTipoValor.EnuInteger)
                    Dim lobjNota As New ClsNotaDb()
                    lobjNota.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentIdNota})
                    lobjNota.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
                    lobjNota.ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuEnProceso
                    lobjNota.SActualice(True)
                Case EnuTipoDocOri.EnuNotaRevCr
                    lstrPref = ClsPanorama.FobjValorCampo(ldrwIdFac(ClsPrefijo_NotaRCrStr.SstrNombreCampoBd),
                            EnuTipoValor.EnuString)
                    lentIdNota = ClsPanorama.FobjValorCampo(ldrwIdFac(ClsIdNotaRCrEnt.SstrNombreCampoBd),
                           EnuTipoValor.EnuInteger)
                    Dim lobjNota As New ClsNotaReversionCr()
                    lobjNota.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentIdNota})
                    lobjNota.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
                    lobjNota.ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuEnProceso
                    lobjNota.SActualice(True)
            End Select
        Next
    End Sub
#End Region
#End Region
#Region "Estado de Cuenta"
#Region "Estado de Cuenta de Clientes"
    Private Sub SGenereEstadosCuentaClientes(adtmFechaEstados As Date)
        Dim ldblIdCliente As Double
        Dim ldtbClientesConDeuda = FdtbClientesConDeuda(True), i = 0
        Dim lobjCliente As New ClsCliente(EnuModoInstanciaObjDef.EnuUnico)
        ObjArgumentoEventoPan.EnuProceso = EnuProcesoDef.EnuGenEstadosCtaCli
        ObjArgumentoEventoPan.DblCantAProcesar = ldtbClientesConDeuda.Rows.Count
        ObjArgumentoEventoPan.DblCantProcesada = i
        RaiseEvent EvnInicio(Me, ObjArgumentoEventoPan)
        RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
        For Each ldrwClieConDeu As DataRow In ldtbClientesConDeuda.Rows
            i += 1
            ldblIdCliente = ClsPanorama.FobjValorCampo(ldrwClieConDeu(
                    ClsIdCliente_FactDbl.SstrNombreCampoBd), EnuTipoValor.EnuDouble)
            lobjCliente.SAbra({GshrIdCarpeta, GshrIdCentroUtil, ldblIdCliente})
            SGenereEstadoCuentaCli(lobjCliente, adtmFechaEstados)
            ObjArgumentoEventoPan.DblCantProcesada = i
            RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
            If ObjArgumentoEventoPan.BlnCancele Then
                Exit Sub
            End If
        Next
    End Sub
    Private Sub SGenereEstadoCuentaCli(aobjCliente As ClsCliente, adtmFechaEstado As Date)
        Dim lobjPrefact As New ClsFactura, lentIdPrefact As Integer
        Dim ldblIdCliente As Double = aobjCliente.ObjIdClienteDbl.ObjValorPro
        Dim ldtbItemsFacVivas = FdtbItemsFacVivas(ldblIdCliente, 0)
        Dim ldtbPrefacts = FdtbIdPrefactsClientes(ldblIdCliente)
        If ldtbPrefacts.Rows.Count > 0 Then
            If aobjCliente.ObjFactPorServicio_CliBln.ObjValorPro Then
                For Each ldrwPrefac As DataRow In ldtbPrefacts.Rows
                    lentIdPrefact = ClsPanorama.FobjValorCampo(ldrwPrefac(0),
                        EnuTipoValor.EnuInteger)
                    lobjPrefact.SAbra({GshrIdCarpeta, GshrIdCentroUtil, GCSTRPREFPREFACTURA,
                        lentIdPrefact})
                    SGenereEstCtaCliePorSer(adtmFechaEstado, ldtbItemsFacVivas,
                            lobjPrefact)
                Next
                For Each ldrwItemFacViva As DataRow In ldtbItemsFacVivas.Rows
                    If ldrwItemFacViva("Proc") = "F" Then
                        SGenereEstCtaClieNoPorSer(aobjCliente, adtmFechaEstado, ldtbItemsFacVivas,
                                "", 0)
                    End If
                Next
            Else
                For Each ldrwprefac As DataRow In ldtbPrefacts.Rows
                    lentIdPrefact = ClsPanorama.FobjValorCampo(ldrwprefac(0),
                            EnuTipoValor.EnuInteger)
                    SGenereEstCtaClieNoPorSer(aobjCliente, adtmFechaEstado,
                            ldtbItemsFacVivas, GCSTRPREFPREFACTURA, lentIdPrefact)
                Next
            End If
        Else
            SGenereEstCtaClieNoPorSer(aobjCliente, adtmFechaEstado, ldtbItemsFacVivas, "", 0)
        End If
    End Sub
    Private Sub SGenereEstCtaCliePorSer(adtmFechaEstado As Date, adtbItemsFacVivas As DataTable,
            aobjFact As ClsFactura)
        Dim lobjFactura As New ClsFactura
        Dim lstrPrefFacViva As String, lentIdFacViva As Integer
        Dim lshrIdAno As Short, lshrIdServ As Short
        Dim ldblIdCliente As Double = aobjFact.ObjIdCliente_FactDbl.ObjValorPro
        Dim lobjEstadoCuenta = FobjNuevoEstadoCuenta(ldblIdCliente, String.Empty,
                adtmFechaEstado)
        For Each ldrwItemFacViva As DataRow In adtbItemsFacVivas.Rows
            If ldrwItemFacViva("Proc") = "F" Then
                lshrIdAno = ClsPanorama.FobjValorCampo(ldrwItemFacViva(
                        ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd),
                        EnuTipoValor.EnuShort)
                If lshrIdAno = 0 Then
                    lshrIdServ = ClsPanorama.FobjValorCampo(ldrwItemFacViva(
                            ClsIdServicio_ItemFactShr.SstrNombreCampoBd),
                            EnuTipoValor.EnuShort)
                    If aobjFact.FblnFacturoServicio(lshrIdAno, lshrIdServ) Then
                        lstrPrefFacViva = ClsPanorama.FobjValorCampo(ldrwItemFacViva(
                                ClsPrefijo_ItemFactStr.SstrNombreCampoBd),
                                EnuTipoValor.EnuString)
                        lentIdFacViva = ClsPanorama.FobjValorCampo(ldrwItemFacViva(
                                ClsIdFactura_ItemFactEnt.SstrNombreCampoBd),
                                EnuTipoValor.EnuInteger)
                        lobjFactura.SAbra({GshrIdCarpeta, GshrIdCentroUtil,
                                lstrPrefFacViva, lentIdFacViva})
                        lobjEstadoCuenta.SAdicioneFacturaEstado(lobjFactura,
                                adtmFechaEstado, lshrIdServ)
                        ldrwItemFacViva("Proc") = "T"
                    End If
                Else
                    ldrwItemFacViva("Proc") = "T"
                End If
            End If
        Next
        If lobjEstadoCuenta.ColFacturasEstado.Count > 0 Then
            lobjEstadoCuenta.ObjPrefijoFac_EstadoStr.ObjValorPro = GCSTRPREFPREFACTURA
            lobjEstadoCuenta.ObjIdFactura_EstadoEnt.ObjValorPro = aobjFact.ObjIdFacturaEnt.ObjValorPro
            lobjEstadoCuenta.SActualice(True)
        End If
    End Sub
    Private Sub SGenereEstCtaClieNoPorSer(aobjClliente As ClsCliente, adtmFechaEstado As Date,
            adtbItemsFacVivas As DataTable, astrPrefijo As String, aentIdPrefac As Integer)
        Dim lobjFactura As New ClsFactura
        Dim lstrPrefFacViva As String, lentIdFacViva As Integer, lstrIdFacPro As String
        Dim lstrIdPreAgr = String.Empty
        Dim lstrFrasProcesadas As New List(Of String)
        Dim ldblIdCliente As Double = aobjClliente.ObjIdClienteDbl.ObjValorPro
        Dim lobjEstadoCuenta = FobjNuevoEstadoCuenta(ldblIdCliente, lstrIdPreAgr,
                adtmFechaEstado)
        For Each ldrwItemVivo As DataRow In adtbItemsFacVivas.Rows
            If ldrwItemVivo("Proc") = "F" Then
                lstrPrefFacViva = ClsPanorama.FobjValorCampo(ldrwItemVivo(
                    ClsPrefijo_ItemFactStr.SstrNombreCampoBd), EnuTipoValor.EnuString)
                lentIdFacViva = ClsPanorama.FobjValorCampo(ldrwItemVivo(
                    ClsIdFactura_ItemFactEnt.SstrNombreCampoBd), EnuTipoValor.EnuInteger)
                lstrIdFacPro = lstrPrefFacViva & "," & lentIdFacViva.ToString()
                If Not lstrFrasProcesadas.Contains(lstrIdFacPro) Then
                    lobjFactura.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefFacViva, lentIdFacViva})
                    lobjEstadoCuenta.SAdicioneFacturaEstado(lobjFactura, adtmFechaEstado)
                    lstrFrasProcesadas.Add(lstrIdFacPro)
                End If
                ldrwItemVivo("Proc") = "T"
            End If
        Next
        If lobjEstadoCuenta.ColFacturasEstado.Count > 0 Then
            lobjEstadoCuenta.ObjPrefijoFac_EstadoStr.ObjValorPro = astrPrefijo
            lobjEstadoCuenta.ObjIdFactura_EstadoEnt.ObjValorPro = aentIdPrefac
            lobjEstadoCuenta.SActualice(True)
        End If
    End Sub
    Private Function FobjEstCtaClieNoPorSer(aobjClliente As ClsCliente, adtmFechaEstado As Date,
            adtbItemsFacVivas As DataTable, astrPrefijo As String, aentIdPrefac As Integer)
        Dim lobjFactura As New ClsFactura
        Dim lstrPrefFacViva As String, lentIdFacViva As Integer, lstrIdFacPro As String
        Dim lstrIdPreAgr = String.Empty
        Dim lstrFrasProcesadas As New List(Of String)
        Dim ldblIdCliente As Double = aobjClliente.ObjIdClienteDbl.ObjValorPro
        Dim lobjEstadoCuenta = FobjNuevoEstadoCuenta(ldblIdCliente, lstrIdPreAgr,
                adtmFechaEstado)
        For Each ldrwItemVivo As DataRow In adtbItemsFacVivas.Rows
            lstrPrefFacViva = ClsPanorama.FobjValorCampo(ldrwItemVivo(
                    ClsPrefijo_ItemFactStr.SstrNombreCampoBd), EnuTipoValor.EnuString)
            lentIdFacViva = ClsPanorama.FobjValorCampo(ldrwItemVivo(
                    ClsIdFactura_ItemFactEnt.SstrNombreCampoBd), EnuTipoValor.EnuInteger)
            lstrIdFacPro = lstrPrefFacViva & "," & lentIdFacViva.ToString()
            If Not lstrFrasProcesadas.Contains(lstrIdFacPro) Then
                lobjFactura.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefFacViva, lentIdFacViva})
                lobjEstadoCuenta.SAdicioneFacturaEstado(lobjFactura, adtmFechaEstado)
                lstrFrasProcesadas.Add(lstrIdFacPro)
            End If
            ldrwItemVivo("Proc") = "T"
        Next
        If lobjEstadoCuenta.ColFacturasEstado.Count > 0 Then
            lobjEstadoCuenta.ObjPrefijoFac_EstadoStr.ObjValorPro = astrPrefijo
            lobjEstadoCuenta.ObjIdFactura_EstadoEnt.ObjValorPro = aentIdPrefac
            lobjEstadoCuenta.SActualice(True)
        End If
        Return lobjEstadoCuenta
    End Function
    Friend Shared Function FobjEstadoCtaHoy(adblIdCliente As Double, astrIdSer As String,
            ByRef adtbItemsFacVivas As DataTable) As ClsEstadoCuenta
        Dim lshrIdServ As Short, lblnTodos As Boolean
        Dim lstrPrefFacViva As String, lentIdFacViva As Integer, lstrIdFacPro As String
        Dim lstrFrasProcesadas As New List(Of String)
        Dim lobjFactura As New ClsFactura
        lblnTodos = astrIdSer = "A"
        lshrIdServ = If(lblnTodos, 0, CType(astrIdSer, Short))
        Dim ldtbItemsFacVivas = FdtbItemsFacVivas(adblIdCliente, lshrIdServ)
        Dim lobjEstadoCuenta = FobjNuevoEstadoCuenta(adblIdCliente, String.Empty,
                Today)
        For Each ldrwItemFacViva As DataRow In ldtbItemsFacVivas.Rows
            If ldrwItemFacViva("Proc") = "F" Then
                lstrPrefFacViva = ClsPanorama.FobjValorCampo(ldrwItemFacViva(
                        ClsPrefijo_ItemFactStr.SstrNombreCampoBd), EnuTipoValor.EnuString)
                lentIdFacViva = ClsPanorama.FobjValorCampo(ldrwItemFacViva(
                        ClsIdFactura_ItemFactEnt.SstrNombreCampoBd),
                        EnuTipoValor.EnuInteger)
                lstrIdFacPro = lstrPrefFacViva & "," & lentIdFacViva.ToString()
                If Not lstrFrasProcesadas.Contains(lstrIdFacPro) Then
                    lstrFrasProcesadas.Add(lstrIdFacPro)
                    lobjFactura.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefFacViva, lentIdFacViva})
                    If lblnTodos Then
                        lobjEstadoCuenta.SAdicioneFacturaEstado(lobjFactura, Today)
                    Else
                        lobjEstadoCuenta.SAdicioneFacturaEstado(lobjFactura, Today, lshrIdServ)
                    End If
                    ldrwItemFacViva("Proc") = "T"
                End If
            Else
                ldrwItemFacViva("Proc") = "T"
            End If
        Next
        adtbItemsFacVivas = ldtbItemsFacVivas
        Return lobjEstadoCuenta
    End Function
    Friend Shared Function FobjEstadoCtaHoy(adblIdCliente As Double, astrIdPredio As String,
            astrIdSer As String, ByRef adtbItemsFacVivas As DataTable) As ClsEstadoCuenta
        Dim lshrIdServ As Short, lblnTodos As Boolean
        Dim lstrPrefFacViva As String, lentIdFacViva As Integer, lstrIdFacPro As String
        Dim lstrFrasProcesadas As New List(Of String)
        Dim lobjFactura As New ClsFactura
        lblnTodos = astrIdSer = "A"
        If Not lblnTodos Then
            lshrIdServ = CShort(astrIdSer)
        End If
        Dim ldtbItemsFacVivas = FdtbItemsFacVivas(adblIdCliente, astrIdPredio, astrIdSer)
        Dim lobjEstadoCuenta = FobjNuevoEstadoCuenta(adblIdCliente, String.Empty,
                Today)
        For Each ldrwItemFacViva As DataRow In ldtbItemsFacVivas.Rows
            If ldrwItemFacViva("Proc") = "F" Then
                lstrPrefFacViva = ClsPanorama.FobjValorCampo(ldrwItemFacViva(
                        ClsPrefijo_ItemFactStr.SstrNombreCampoBd),
                        EnuTipoValor.EnuString)
                lentIdFacViva = ClsPanorama.FobjValorCampo(ldrwItemFacViva(
                        ClsIdFactura_ItemFactEnt.SstrNombreCampoBd),
                        EnuTipoValor.EnuInteger)
                lstrIdFacPro = lstrPrefFacViva & "," & lentIdFacViva.ToString()
                If Not lstrFrasProcesadas.Contains(lstrIdFacPro) Then
                    lobjFactura.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefFacViva,
                            lentIdFacViva})
                    If lblnTodos Then
                        lobjEstadoCuenta.SAdicioneFacturaEstado(lobjFactura, Today)
                    Else
                        lobjEstadoCuenta.SAdicioneFacturaEstado(lobjFactura, Today,
                            lshrIdServ)
                    End If
                    lstrFrasProcesadas.Add(lstrIdFacPro)
                End If
                ldrwItemFacViva("Proc") = "T"
            End If
        Next
        adtbItemsFacVivas = ldtbItemsFacVivas
        Return lobjEstadoCuenta
    End Function
#End Region
#Region "Estado de Cuenta de Predios"
    Private Sub SGenereEstadoCtaPrediosAgr(adtmFechaEstados As Date)
        Dim lstrIdPreAgr As String, i = 0
        Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.EnuUnico)
        Dim ldtbPrediosConDeuda = FdtbPrediosAgrConDeuda()
        ObjArgumentoEventoPan.EnuProceso = EnuProcesoDef.EnuGenEstadosCtaPre
        ObjArgumentoEventoPan.DblCantAProcesar = ldtbPrediosConDeuda.Rows.Count
        ObjArgumentoEventoPan.DblCantProcesada = i
        RaiseEvent EvnInicio(Me, ObjArgumentoEventoPan)
        RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
        For Each ldrwPreConDeu As DataRow In ldtbPrediosConDeuda.Rows
            i += 1
            lstrIdPreAgr = ClsPanorama.FobjValorCampo(ldrwPreConDeu(
                    ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd),
                    EnuTipoValor.EnuString)
            lobjPredio.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrIdPreAgr})
            SGenereEstadoCuentaPredAgr(lobjPredio, adtmFechaEstados)
            ObjArgumentoEventoPan.DblCantProcesada = i
            RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
            If ObjArgumentoEventoPan.BlnCancele Then
                Exit Sub
            End If
        Next
    End Sub
    Private Sub SGenereEstadoCuentaPredAgr(aobjPredio As ClsPredio, adtmFechaEstado As Date)
        Dim ldblIdCliente As Double
        Dim lobjPrefact As New ClsFactura, lentIdPrefact As Integer
        Dim lstrIdPredAgr As String = aobjPredio.ObjIdPredioStr.ObjValorPro
        Dim ldtbItemsFacVivas = FdtbItemsFacVivasPreAgr(lstrIdPredAgr, 9999, 0)
        Dim ldtbPrefacts = FdtbIdPrefactsPredAgr(lstrIdPredAgr)
        If ldtbPrefacts.Rows.Count > 0 Then
            For Each ldrwPrefac As DataRow In ldtbPrefacts.Rows
                lentIdPrefact = ClsPanorama.FobjValorCampo(ldrwPrefac(0),
                    EnuTipoValor.EnuInteger)
                lobjPrefact.SAbra({GshrIdCarpeta, GshrIdCentroUtil, GCSTRPREFPREFACTURA,
                        lentIdPrefact})
                ldblIdCliente = lobjPrefact.ObjIdCliente_FactDbl.ObjValorPro
                lstrIdPredAgr = lobjPrefact.ObjIdPredioAgrupador_FacStr.ObjValorPro
                If aobjPredio.ObjFacturarPorServicio_PreBln.ObjValorPro Then
                    SGenereEstCtaPreAgrPorSer(lobjPrefact, ldtbItemsFacVivas, adtmFechaEstado)
                Else
                    SGenereEstCtaPredAgrNoPorSer(lobjPrefact, ldtbItemsFacVivas, adtmFechaEstado)
                End If
            Next
        End If
        SGenereEstaCtaPredSinFac(aobjPredio, ldtbItemsFacVivas, adtmFechaEstado)
    End Sub
    Private Sub SGenereEstCtaPreAgrPorSer(aobjPreFactura As ClsFactura, adtbFactVivas As DataTable,
            adtmFechaEstado As Date)
        Dim lobjFactViva As New ClsFactura
        Dim lstrPrefFacViva As String, lentIdFacViva As Integer
        Dim lshrIdAno As Short, lshrIdServ As Short
        Dim ldblIdCliente As Double = aobjPreFactura.ObjIdCliente_FactDbl.ObjValorPro
        Dim lstrIdPredAgr As String = aobjPreFactura.ObjIdPredioAgrupador_FacStr.ObjValorPro
        Dim lobjEstadoCuenta = FobjNuevoEstadoCuenta(ldblIdCliente, lstrIdPredAgr,
                adtmFechaEstado)
        For Each ldrwItemFacViva As DataRow In adtbFactVivas.Rows
            If ldrwItemFacViva("Proc") = "F" Then
                lshrIdAno = ClsPanorama.FobjValorCampo(ldrwItemFacViva(
                        ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd), EnuTipoValor.EnuShort)
                lshrIdServ = ClsPanorama.FobjValorCampo(ldrwItemFacViva(
                    ClsIdServicio_ItemFactShr.SstrNombreCampoBd), EnuTipoValor.EnuShort)
                If aobjPreFactura.FblnFacturoServicio(lshrIdAno, lshrIdServ) OrElse
                        lshrIdAno > 0 Then
                    lstrPrefFacViva = ClsPanorama.FobjValorCampo(ldrwItemFacViva(
                            ClsPrefijo_ItemFactStr.SstrNombreCampoBd),
                            EnuTipoValor.EnuString)
                    lentIdFacViva = ClsPanorama.FobjValorCampo(ldrwItemFacViva(
                            ClsIdFactura_ItemFactEnt.SstrNombreCampoBd),
                            EnuTipoValor.EnuInteger)
                    lobjFactViva.SAbra({GshrIdCarpeta, GshrIdCentroUtil,
                            lstrPrefFacViva, lentIdFacViva})
                    lobjEstadoCuenta.SAdicioneFacturaEstado(lobjFactViva, adtmFechaEstado,
                            lshrIdServ)
                    ldrwItemFacViva("Proc") = "T"
                End If
            End If
        Next
        If lobjEstadoCuenta.ColFacturasEstado.Count > 0 Then
            lobjEstadoCuenta.ObjPrefijoFac_EstadoStr.ObjValorPro = GCSTRPREFPREFACTURA
            lobjEstadoCuenta.ObjIdFactura_EstadoEnt.ObjValorPro =
                    aobjPreFactura.ObjIdFacturaEnt.ObjValorPro
            lobjEstadoCuenta.SActualice(True)
        End If
    End Sub
    Private Sub SGenereEstCtaPredAgrNoPorSer(aobjPrefact As ClsFactura, adtbItemsFacVivas As DataTable,
                adtmFechaEstado As Date)
        Dim lobjFactViva As New ClsFactura
        Dim lstrPrefFacViva As String, lentIdFacViva As Integer, lstrIdFacPro As String
        Dim lstrFrasProcesadas As New List(Of String)
        Dim ldblIdCliente As Double = aobjPrefact.ObjIdCliente_FactDbl.ObjValorPro
        Dim lstrIdPredAgr As String = aobjPrefact.ObjIdPredioAgrupador_FacStr.ObjValorPro
        Dim lobjEstadoCuenta = FobjNuevoEstadoCuenta(ldblIdCliente, lstrIdPredAgr, adtmFechaEstado)
        For Each ldrwItemVivo As DataRow In adtbItemsFacVivas.Rows
            If ldrwItemVivo("Proc") = "F" Then
                lstrPrefFacViva = ClsPanorama.FobjValorCampo(ldrwItemVivo(
                        ClsPrefijo_ItemFactStr.SstrNombreCampoBd), EnuTipoValor.EnuString)
                lentIdFacViva = ClsPanorama.FobjValorCampo(ldrwItemVivo(
                    ClsIdFactura_ItemFactEnt.SstrNombreCampoBd), EnuTipoValor.EnuInteger)
                lstrIdFacPro = lstrPrefFacViva & "," & lentIdFacViva.ToString()
                If Not lstrFrasProcesadas.Contains(lstrIdFacPro) Then
                    lobjFactViva.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefFacViva,
                            lentIdFacViva})
                    If lobjFactViva.ObjIdCliente_FactDbl.ObjValorPro = ldblIdCliente Then
                        lobjEstadoCuenta.SAdicioneFacturaEstado(lobjFactViva, adtmFechaEstado)
                        ldrwItemVivo("Proc") = "T"
                    End If
                    lstrFrasProcesadas.Add(lstrIdFacPro)
                End If
            End If
        Next
        If lobjEstadoCuenta.ColFacturasEstado.Count > 0 Then
            lobjEstadoCuenta.ObjPrefijoFac_EstadoStr.ObjValorPro =
                    aobjPrefact.ObjPrefijo_FactStr.ObjValorPro
            lobjEstadoCuenta.ObjIdFactura_EstadoEnt.ObjValorPro =
                    aobjPrefact.ObjIdFacturaEnt.ObjValorPro
            lobjEstadoCuenta.SActualice(True)
        End If
    End Sub
    ''' <summary>
    ''' Genera el estado de cuenta para toas aquella deudas del predio agrupador a las cuales
    ''' no les corresponde una prefactura
    ''' </summary>
    Private Sub SGenereEstaCtaPredSinFac(aobjPredio As ClsPredio, adtbItemsFacVivas As DataTable,
            adtmFechaEstado As Date)
        Dim lstrIdPredAgr As String = aobjPredio.ObjIdPredioAgrupadorStr.ObjValorPro
        Dim lobjFactViva As New ClsFactura, lobjEstadoCuenta As ClsEstadoCuenta = Nothing
        Dim lstrFiltro = "Proc = 'F'", lstrPrefFacViva As String, lentIdFacViva As Integer
        Dim ldblIdClienteFac As Double = 0
        Dim ldrwItemsFacVivas As DataRow() = adtbItemsFacVivas.Select(lstrFiltro)
        For Each ldrwItemVivo As DataRow In ldrwItemsFacVivas
            If ldrwItemVivo("Proc") = "F" Then
                lstrPrefFacViva = ClsPanorama.FobjValorCampo(ldrwItemVivo(
                    ClsPrefijo_ItemFactStr.SstrNombreCampoBd), EnuTipoValor.EnuString)
                lentIdFacViva = ClsPanorama.FobjValorCampo(ldrwItemVivo(
                        ClsIdFactura_ItemFactEnt.SstrNombreCampoBd), EnuTipoValor.EnuInteger)
                lobjFactViva.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefFacViva, lentIdFacViva})
                If ldblIdClienteFac = lobjFactViva.ObjIdCliente_FactDbl.ObjValorNuevo Then
                    lobjEstadoCuenta.SAdicioneFacturaEstado(lobjFactViva, adtmFechaEstado)
                    ldrwItemVivo("Proc") = "T"
                Else
                    If lobjEstadoCuenta IsNot Nothing AndAlso
                            lobjEstadoCuenta.ColFacturasEstado.Count > 0 Then
                        lobjEstadoCuenta.ObjPrefijoFac_EstadoStr.ObjValorPro = ""
                        lobjEstadoCuenta.ObjIdFactura_EstadoEnt.ObjValorPro = 0
                        lobjEstadoCuenta.SActualice(True)
                    End If
                    ldblIdClienteFac = lobjFactViva.ObjIdCliente_FactDbl.ObjValorNuevo
                    lobjEstadoCuenta = FobjNuevoEstadoCuenta(ldblIdClienteFac, lstrIdPredAgr,
                            adtmFechaEstado)
                    lobjEstadoCuenta.SAdicioneFacturaEstado(lobjFactViva, adtmFechaEstado)
                    ldrwItemVivo("Proc") = "T"
                End If
            End If
        Next
        If lobjEstadoCuenta IsNot Nothing Then
            lobjEstadoCuenta.ObjPrefijoFac_EstadoStr.ObjValorPro = ""
            lobjEstadoCuenta.ObjIdFactura_EstadoEnt.ObjValorPro = 0
            lobjEstadoCuenta.SActualice(True)
        End If
    End Sub
    Private Sub SGenereEstCtaPredNoPorSer(astrIdPredAgru As String, adtmFechaEstado As Date,
            adtbItemsFacVivas As DataTable, astrPrefijo As String, aentIdFact As Integer)
        Dim lobjFactViva As New ClsFactura
        Dim lstrPrefFacViva As String, lentIdFacViva As Integer, lstrIdFacPro As String
        Dim lstrFrasProcesadas As New List(Of String)
        Dim lobjEstadoCuenta = FobjNuevoEstadoCuenta(0, astrIdPredAgru,
                adtmFechaEstado)
        For Each ldrwItemVivo As DataRow In adtbItemsFacVivas.Rows
            If ldrwItemVivo("Proc") = "F" Then
                lstrPrefFacViva = ClsPanorama.FobjValorCampo(ldrwItemVivo(
                        ClsPrefijo_ItemFactStr.SstrNombreCampoBd), EnuTipoValor.EnuString)
                lentIdFacViva = ClsPanorama.FobjValorCampo(ldrwItemVivo(
                        ClsIdFactura_ItemFactEnt.SstrNombreCampoBd), EnuTipoValor.EnuInteger)
                lstrIdFacPro = lstrPrefFacViva & "," & lentIdFacViva.ToString()
                If Not lstrFrasProcesadas.Contains(lstrIdFacPro) Then
                    lobjFactViva.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefFacViva, lentIdFacViva})
                    lobjEstadoCuenta.SAdicioneFacturaEstado(lobjFactViva, adtmFechaEstado)
                    lstrFrasProcesadas.Add(lstrIdFacPro)
                End If
                ldrwItemVivo("Proc") = "T"
            End If
        Next
        If lobjEstadoCuenta.ColFacturasEstado.Count > 0 Then
            lobjEstadoCuenta.ObjPrefijoFac_EstadoStr.ObjValorPro = astrPrefijo
            lobjEstadoCuenta.ObjIdFactura_EstadoEnt.ObjValorPro = aentIdFact
            lobjEstadoCuenta.SActualice(True)
        End If
    End Sub
    Private Function FobjEstCtaPredNoPorSer(adblIdCliente As Double, astrIdPredAgru As String,
            adtmFechaEstado As Date, adtbItemsFacVivas As DataTable, astrPrefijo As String,
            aentIdFact As Integer) As ClsEstadoCuenta
        Dim lobjFactura As New ClsFactura
        Dim lstrPrefFacViva As String, lentIdFacViva As Integer, lstrIdFacPro As String
        Dim lstrFrasProcesadas As New List(Of String)
        Dim lobjEstadoCuenta = FobjNuevoEstadoCuenta(adblIdCliente, astrIdPredAgru,
                adtmFechaEstado)
        For Each ldrwItemVivo As DataRow In adtbItemsFacVivas.Rows
            lstrPrefFacViva = ClsPanorama.FobjValorCampo(ldrwItemVivo(
                    ClsPrefijo_ItemFactStr.SstrNombreCampoBd), EnuTipoValor.EnuString)
            lentIdFacViva = ClsPanorama.FobjValorCampo(ldrwItemVivo(
                    ClsIdFactura_ItemFactEnt.SstrNombreCampoBd), EnuTipoValor.EnuInteger)
            lstrIdFacPro = lstrPrefFacViva & "," & lentIdFacViva.ToString()
            If Not lstrFrasProcesadas.Contains(lstrIdFacPro) Then
                lobjFactura.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefFacViva, lentIdFacViva})
                lobjEstadoCuenta.SAdicioneFacturaEstado(lobjFactura, adtmFechaEstado)
                lstrFrasProcesadas.Add(lstrIdFacPro)
            End If
            ldrwItemVivo("Proc") = "T"
        Next
        If lobjEstadoCuenta.ColFacturasEstado.Count > 0 Then
            lobjEstadoCuenta.ObjPrefijoFac_EstadoStr.ObjValorPro = astrPrefijo
            lobjEstadoCuenta.ObjIdFactura_EstadoEnt.ObjValorPro = aentIdFact
            lobjEstadoCuenta.SActualice(True)
        End If
        Return lobjEstadoCuenta
    End Function
    Private Sub SGenereEstCtaAPropYPreAgr(adtmFechaEstado As Date,
            adtbItemsFacVivas As DataTable, aobjFact As ClsFactura)
        Dim lstrPrefFacViva As String, lentIdFacViva As Integer, lstrIdFacPro As String
        Dim lstrFrasProcesadas As New List(Of String), lobjFactura As New ClsFactura
        Dim lshrIdSerFacViva As Short, lshrIdSerFact As Short
        Dim lobjItemfac As ClsItemFactura = aobjFact.ColItemsFactura("1")
        Dim lstrIdpredAgr As String = aobjFact.ObjIdPredioAgrupador_FacStr.ObjValorPro
        Dim lentIdFact As Integer = aobjFact.ObjIdFacturaEnt.ObjValorPro
        Dim lstrPref As String = aobjFact.ObjPrefijo_FactStr.ObjValorPro
        lshrIdSerFact = lobjItemfac.ObjIdServicio_ItemFactShr.ObjValorPro
        Dim lobjEstadoCuenta = FobjNuevoEstadoCuenta(0, lstrIdpredAgr,
                adtmFechaEstado)
        For Each ldrwItemVivo As DataRow In adtbItemsFacVivas.Rows
            lstrPrefFacViva = ClsPanorama.FobjValorCampo(ldrwItemVivo(
                    ClsPrefijo_ItemFactStr.SstrNombreCampoBd), EnuTipoValor.EnuString)
            lentIdFacViva = ClsPanorama.FobjValorCampo(ldrwItemVivo(
                    ClsIdFactura_ItemFactEnt.SstrNombreCampoBd),
                    EnuTipoValor.EnuInteger)
            lstrIdFacPro = lstrPrefFacViva & "," & lentIdFacViva.ToString()
            If Not lstrFrasProcesadas.Contains(lstrIdFacPro) Then
                lobjFactura.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefFacViva,
                        lentIdFacViva})
                If lobjFactura.BlnFacturaAPropYPreAgr Then
                    lobjItemfac = lobjFactura.ColItemsFactura("1")
                    lshrIdSerFacViva = lobjItemfac.ObjIdServicio_ItemFactShr.ObjValorPro
                    If lshrIdSerFact = lshrIdSerFacViva Then
                        lobjEstadoCuenta.SAdicioneFacturaEstado(lobjFactura,
                        adtmFechaEstado)
                        ldrwItemVivo("Proc") = "T"
                    End If
                Else
                    ldrwItemVivo("Proc") = "T"
                End If
                lstrFrasProcesadas.Add(lstrIdFacPro)
            Else
                ldrwItemVivo("Proc") = "T"
            End If
        Next
        If lobjEstadoCuenta.ColFacturasEstado.Count > 0 Then
            lobjEstadoCuenta.ObjPrefijoFac_EstadoStr.ObjValorPro = lstrPref
            lobjEstadoCuenta.ObjIdFactura_EstadoEnt.ObjValorPro = lentIdFact
            lobjEstadoCuenta.SActualice(True)
        End If
    End Sub
    Private Sub SGenereEstCtaPredPorSer(adtmFechaEstado As Date,
            adtbItemsFacVivas As DataTable, aobjPrefac As ClsFactura)
        Dim lobjFactura As New ClsFactura
        Dim lstrPrefFacViva As String, lentIdFacViva As Integer
        Dim lshrIdAno As Short, lshrIdServ As Short
        Dim lstrIdpredAgr As String = aobjPrefac.ObjIdPredioAgrupador_FacStr.ObjValorPro
        Dim lobjEstadoCuenta = FobjNuevoEstadoCuenta(0, lstrIdpredAgr, adtmFechaEstado)
        For Each ldrwItemFacViva As DataRow In adtbItemsFacVivas.Rows
            If ldrwItemFacViva("Proc") = "F" Then
                lshrIdAno = ClsPanorama.FobjValorCampo(ldrwItemFacViva(
                        ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd),
                        EnuTipoValor.EnuShort)
                If lshrIdAno = 0 Then
                    lshrIdServ = ClsPanorama.FobjValorCampo(ldrwItemFacViva(
                            ClsIdServicio_ItemFactShr.SstrNombreCampoBd),
                            EnuTipoValor.EnuShort)
                    If aobjPrefac.FblnFacturoServicio(lshrIdAno, lshrIdServ) Then
                        lstrPrefFacViva = ClsPanorama.FobjValorCampo(ldrwItemFacViva(
                                ClsPrefijo_ItemFactStr.SstrNombreCampoBd),
                                EnuTipoValor.EnuString)
                        lentIdFacViva = ClsPanorama.FobjValorCampo(ldrwItemFacViva(
                                ClsIdFactura_ItemFactEnt.SstrNombreCampoBd),
                                EnuTipoValor.EnuInteger)
                        lobjFactura.SAbra({GshrIdCarpeta, GshrIdCentroUtil,
                                lstrPrefFacViva, lentIdFacViva})
                        lobjEstadoCuenta.SAdicioneFacturaEstado(lobjFactura,
                                adtmFechaEstado, lshrIdServ)
                        ldrwItemFacViva("Proc") = "T"
                    End If
                End If
            End If
        Next
        If lobjEstadoCuenta.ColFacturasEstado.Count > 0 Then
            lobjEstadoCuenta.ObjPrefijoFac_EstadoStr.ObjValorPro = GCSTRPREFPREFACTURA
            lobjEstadoCuenta.ObjIdFactura_EstadoEnt.ObjValorPro = aobjPrefac.ObjIdFacturaEnt.ObjValorPro
            lobjEstadoCuenta.SActualice(True)
        End If
    End Sub
    Private Shared Function FdtbPrediosAgrConDeuda() As DataTable
        Dim lstrTabla = ClsFactura.SstrNombreTabla
        Dim lstrCampSele = {"DISTINCT " & ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd}
        ' Filtro
        Dim lstrFiltro = StrFiltroUbicacion & " AND " &
                ClsDebitos_FactDec.SstrNombreCampoBd & " <> " &
                ClsCreditos_FactDec.SstrNombreCampoBd & " AND " &
                ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd & " <> ''" & " AND " &
                ClsPrefijo_FactStr.SstrNombreCampoBd & " <> '" &
                GCSTRPREFPREFACTURA & "'"
        Dim lstrOrden(,) As String = {{ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd, "ASC"},
                {ClsIdCliente_FactDbl.SstrNombreCampoBd, "ASC"}}
        Dim ldtbPrediosConDeuda = ClsPanorama.FdtbDataTable(lstrTabla, lstrCampSele,
                lstrOrden, lstrFiltro, False, Array.Empty(Of String))
        Return ldtbPrediosConDeuda
    End Function
    Friend Shared Function FdtbPrediosAgrMorosos(aentDiasVencido As Integer,
            adblUltimoClientenviado As Double) As DataTable
        Dim ldtmFechaEstaVen = Date.Today.AddDays(-aentDiasVencido)
        Dim lstrFechaEstaVencida = "'" & ClsPanoramaDat.FstrFechaNormalizada(
                ldtmFechaEstaVen) & "'"
        Dim lstrTabla = ClsFactura.SstrNombreTabla
        Dim lstrCampSele = {"DISTINCT " & ClsIdCliente_FactDbl.SstrNombreCampoBd,
                ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd}
        ' Filtro
        Dim lstrFiltro = StrFiltroUbicacion & " AND " &
                ClsIdCliente_FactDbl.SstrNombreCampoBd & " > " & adblUltimoClientenviado & " AND " &
                ClsPrefijo_FactStr.SstrNombreCampoBd & " <> '" &
                GCSTRPREFPREFACTURA & "' AND " &
                ClsFechaVencimientoDtm.SstrNombreCampoBd & " < " & lstrFechaEstaVencida &
                " AND " & ClsDebitos_FactDec.SstrNombreCampoBd & " <> " &
                ClsCreditos_FactDec.SstrNombreCampoBd & " AND " &
                ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd & " <> ''"
        Dim lstrOrden As String(,) = {{ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd,
                "ASC"}}
        Dim ldtbPrediosMorosos = ClsPanorama.FdtbDataTable(lstrTabla, lstrCampSele,
                lstrOrden, lstrFiltro, False, Array.Empty(Of String))
        Return ldtbPrediosMorosos
    End Function
#End Region
#Region "Estados de Cuenta para Prefacturas"
    Private Sub SGenereEstadosCtaPrefacturas()
        Dim lobjPreFactura As New ClsFactura(GCSTRPREFPREFACTURA)
        Dim lblnFactAPropYPreAgr As Boolean, lblnFactPorSer As Boolean, i = 0.0
        Dim ldtbPrefacturas = FdtbPrefacturas(), ldtbItemsFacVivas As DataTable
        Dim ldblCantFrasAProcesar = FdblCantidadPreFacAProcesar(ldtbPrefacturas)
        Dim lentIdFacMax = 0
        If GobjParametros.BlnEFacAutorizado Then
            lentIdFacMax = GobjParametros.ObjRangoFraFinEnt.ObjValorPro
        End If
        ObjArgumentoEventoPan.EnuProceso = EnuProcesoDef.EnuGenEstadosCtaPre
        ObjArgumentoEventoPan.DblCantAProcesar = ldblCantFrasAProcesar
        RaiseEvent EvnInicio(Me, ObjArgumentoEventoPan)
        lobjPreFactura.SVayaAlPrimero()
        Do While lobjPreFactura.BlnExiste
            i += 1
            Dim lentIdFactura As Integer = lobjPreFactura.ObjIdFacturaEnt.ObjValorPro
            If lentIdFacMax = 0 OrElse lentIdFactura <= lentIdFacMax Then
                Dim ldtmFechaEst As Date = lobjPreFactura.ObjFechaFacturaDtm.ObjValorPro.AddDays(-1)
                Dim lstrIdPreAgr As String = lobjPreFactura.ObjIdPredioAgrupador_FacStr.ToString()
                Dim ldblIdCliente As Double = lobjPreFactura.ObjIdCliente_FactDbl.ObjValorPro
                lblnFactPorSer = If(lobjPreFactura.ObjPredioAgrFactura IsNot Nothing,
                        lobjPreFactura.ObjPredioAgrFactura.ObjFacturarPorServicio_PreBln.ObjValorPro,
                        lobjPreFactura.ObjClienteFactura.ObjFactPorServicio_CliBln.ObjValorPro)
                Dim lobjItemFac As ClsItemFactura = lobjPreFactura.ColItemsFactura("1")
                lblnFactAPropYPreAgr = lobjItemFac.ObjServicio.ObjFactAPropYPreAgrBln.ObjValorPro
                If Not String.IsNullOrEmpty(lstrIdPreAgr) Then
                    ldtbItemsFacVivas = FdtbItemsFacVivas(ldblIdCliente, lstrIdPreAgr, "A")
                    If lblnFactAPropYPreAgr Then
                        SGenereEstCtaAPropYPreAgr(ldtmFechaEst, ldtbItemsFacVivas,
                                lobjPreFactura)
                    ElseIf lblnFactPorSer Then
                        SGenereEstCtaPredPorSer(ldtmFechaEst, ldtbItemsFacVivas,
                                lobjPreFactura)
                    Else
                        SGenereEstCtaPredNoPorSer(lstrIdPreAgr, ldtmFechaEst,
                                ldtbItemsFacVivas, GCSTRPREFPREFACTURA,
                                lobjPreFactura.ObjIdFacturaEnt.ObjValorPro)
                    End If
                Else
                    ldtbItemsFacVivas = FdtbItemsFacVivas(ldblIdCliente, 0)
                    If lblnFactPorSer Then
                        SGenereEstCtaCliePorSer(ldtmFechaEst, ldtbItemsFacVivas,
                                lobjPreFactura)
                    Else
                        Dim lobjCliente As New ClsCliente
                        lobjCliente.SAbra({GshrIdCarpeta, GshrIdCentroUtil,
                                ldblIdCliente})
                        SGenereEstCtaClieNoPorSer(lobjCliente, ldtmFechaEst,
                                ldtbItemsFacVivas, GCSTRPREFPREFACTURA,
                                lobjPreFactura.ObjIdFacturaEnt.ObjValorPro)
                    End If
                End If
            End If
            ObjArgumentoEventoPan.DblCantProcesada = i
            RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
            If ObjArgumentoEventoPan.BlnCancele Then
                Exit Do
            End If
            lobjPreFactura.SVayaAlSiguiente()
        Loop
    End Sub
    ''' <summary>
    ''' Devuelve la cantida de Prefacturas a facturar teniendo en cuenta la reolución de la DIAN
    ''' </summary>
    ''' <returns></returns>
    Private Function FdblCantidadPreFacAProcesar(adtbPrefacturas As DataTable) As Double
        Dim ldblCantAProcesar As Double
        If GobjParametros.BlnEFacAutorizado Then
            Dim lentIdFacMax = GobjParametros.ObjRangoFraFinEnt.ObjValorPro
            Dim lentUltReg = adtbPrefacturas.Rows.Count - 1
            Dim ldrwUltReg = adtbPrefacturas.Rows(lentUltReg)
            Dim lstrNomCampo = ClsIdFacturaEnt.SstrNombreCampoBd
            Dim lentIdUltimaFac = ClsPanorama.FobjValorCampo(ldrwUltReg(lstrNomCampo),
EnuTipoValor.EnuInteger)
            If lentIdFacMax > lentIdUltimaFac Then
                ldblCantAProcesar = adtbPrefacturas.Rows.Count
            Else
                Dim lentIdPrimeraFac = adtbPrefacturas.Rows(0)(lstrNomCampo)
                ldblCantAProcesar = lentIdFacMax - lentIdPrimeraFac + 1
            End If
        Else
            ldblCantAProcesar = adtbPrefacturas.Rows.Count
        End If
        Return ldblCantAProcesar
    End Function
#End Region
    Private Sub SGenereEstadosCtaFactCont(adtbFactCont As DataTable)
        Dim lobjFactura As New ClsFactura()
        Dim lobjCliente As New ClsCliente(EnuModoInstanciaObjDef.EnuUnico)
        Dim lblnFactAPropYPreAgr As Boolean, lblnFactPorSer As Boolean
        Dim ldtbItemsFacViva As DataTable
        Dim ldblIdCliente As Double, lstrIdPreAgr As String, ldtmFechaEst As Date
        For Each ldrwItemFac As DataRow In adtbFactCont.Rows
            Dim lstrPrefFac As String = ClsPanorama.FobjValorCampo(
                    ldrwItemFac("PrefFactura"), EnuTipoValor.EnuString)
            Dim lentIdFac As Integer = ClsPanorama.FobjValorCampo(
                    ldrwItemFac("IdFactura"), EnuTipoValor.EnuInteger)
            lobjFactura.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefFac, lentIdFac})
            ldtmFechaEst = lobjFactura.ObjFechaFacturaDtm.ObjValorPro.AddDays(-1)
            lstrIdPreAgr = lobjFactura.ObjIdPredioAgrupador_FacStr.ToString()
            ldblIdCliente = lobjFactura.ObjIdCliente_FactDbl.ObjValorPro
            With lobjFactura
                lblnFactPorSer = If(.ObjPredioAgrFactura IsNot Nothing,
                        .ObjPredioAgrFactura.ObjFacturarPorServicio_PreBln.ObjValorPro,
                        .ObjClienteFactura.ObjFactPorServicio_CliBln.ObjValorPro)
            End With
            Dim lobjItemFac As ClsItemFactura = lobjFactura.ColItemsFactura("1")
            lblnFactAPropYPreAgr =
                    lobjItemFac.ObjServicio.ObjFactAPropYPreAgrBln.ObjValorPro
            If Not String.IsNullOrEmpty(lstrIdPreAgr) Then
                ldtbItemsFacViva = FdtbItemsFacVivas(ldblIdCliente, lstrIdPreAgr, "A")
                If lblnFactAPropYPreAgr Then
                    SGenereEstCtaAPropYPreAgr(ldtmFechaEst, ldtbItemsFacViva, lobjFactura)
                ElseIf lblnFactPorSer Then
                    SGenereEstCtaPredPorSer(ldtmFechaEst, ldtbItemsFacViva, lobjFactura)
                Else
                    SGenereEstCtaPredNoPorSer(lstrIdPreAgr, ldtmFechaEst,
                            ldtbItemsFacViva, lstrPrefFac, lentIdFac)
                End If
            Else
                ldtbItemsFacViva = FdtbItemsFacVivas(ldblIdCliente, 0)
                If lblnFactPorSer Then
                    SGenereEstCtaCliePorSer(ldtmFechaEst, ldtbItemsFacViva, lobjFactura)
                Else
                    lobjCliente.SAbra({GshrIdCarpeta, GshrIdCentroUtil, ldblIdCliente})
                    SGenereEstCtaClieNoPorSer(lobjCliente, ldtmFechaEst,
                            ldtbItemsFacViva, lstrPrefFac, lentIdFac)
                End If
            End If
        Next
    End Sub
    ''' <summary>
    ''' Devuelve un estado de cuenta según los parametros pasados en los argumentos
    ''' </summary>
    ''' <param name="adblIdCliente">Si adblIdCliente = 0 entonces es estado de 
    ''' cuenta es para un predio agrupdor diferente de ninguno</param>
    ''' <returns></returns>
    Private Shared Function FobjNuevoEstadoCuenta(adblIdCliente As Double,
            astrIdPredioAgru As String, adtmFechaEstado As Date) As ClsEstadoCuenta
        Dim ldrwNuevoEstadoCta As DataRow = FdrwNuevoEstadoCta()
        Dim lobjEstadoCuenta As New ClsEstadoCuenta(ldrwNuevoEstadoCta)
        With lobjEstadoCuenta
            If Not CType(.EnuPermisosObj And EnuPermisosDef.EnuCrear, Boolean) Then
                .EnuPermisosObj += EnuPermisosDef.EnuCrear
            End If
            .SCreeObj(Nothing)
            .ObjDeudaCapitalDec.ObjValorPro = 0
            .ObjDeudaIntMoraDec.ObjValorPro = 0
            .ObjFechaEstadoDtm.ObjValorPro = adtmFechaEstado
            .ObjIdCarpeta_EstadoShr.ObjValorPro = GshrIdCarpeta
            .ObjIdCentroUtil_EstadoShr.ObjValorPro = GshrIdCentroUtil
            .ObjIdPredioAgr_EstadoStr.ObjValorPro = astrIdPredioAgru
            .ObjIdCliente_EstadoDbl.ObjValorPro = adblIdCliente
            .ObjPrefijoFac_EstadoStr.ObjValorPro = String.Empty
            .ObjIdFactura_EstadoEnt.ObjValorPro = 0
            .ObjAntPorAplDec.ObjValorPro = 0
        End With
        Return lobjEstadoCuenta
    End Function
    Private Shared Function FdtbClientesConDeuda(ablnSinPredioAgr As Boolean) As DataTable
        Dim lstbClieConDeuda As New StringBuilder
        Dim lstrCamposSelect = {"DISTINCT " & ClsIdCliente_FactDbl.SstrNombreCampoBd}
        ' Filtro
        With lstbClieConDeuda
            .Clear().Append(StrFiltroUbicacion)
            If ablnSinPredioAgr Then
                .Append(" AND ").Append(ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd)
                .Append(" = ''").Append(" AND ").Append(ClsPrefijo_FactStr.SstrNombreCampoBd)
                .Append(" <> '").Append(GCSTRPREFPREFACTURA).Append("'")
            End If
            .Append(" AND ").Append(ClsDebitos_FactDec.SstrNombreCampoBd)
            .Append(" <> ").Append(ClsCreditos_FactDec.SstrNombreCampoBd)
        End With
        Dim lstrFiltro = lstbClieConDeuda.ToString
        ' Orden
        Dim lstrOrden(,) As String = {{ClsIdCliente_FactDbl.SstrNombreCampoBd, "ASC"}}
        Dim ldtbClientesConDeuda = ClsPanorama.FdtbDataTable(ClsFactura.SstrNombreTabla,
                lstrCamposSelect, lstrOrden, lstrFiltro)
        Return ldtbClientesConDeuda
    End Function
    ''' <summary>
    ''' Devuelve un datatable que contienen las facturas vivas. 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>    
    Private Shared Function FdrwNuevoEstadoCta() As DataRow
        Dim lstrIndice = {{StrCampoCarpeta, "ASC"},
                {StrCampoCentroUtil, "ASC"},
                {ClsIdEstadoCuentaEnt.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " &
                ClsIdEstadoCuentaEnt.SstrNombreCampoBd & " = 0"
        Dim lstrCamposSelect() = {"*"}
        Dim ldtbEstadoCuenta = ClsPanorama.FdtbDataTable(ClsEstadoCuenta.SstrNombreTabla,
                lstrCamposSelect, lstrIndice, lstrFiltro)
        Return ldtbEstadoCuenta.NewRow
    End Function
    Private Shared Function FdtmFechaPrefacturas() As Date
        Dim ldtmFecha = GCDTMFECHANULA
        Dim lstrPrefPrefacturas = GCSTRPREFPREFACTURA
        Dim lstrCamposSelect = {"DISTINCT " & ClsFechaFacturaDtm.SstrNombreCampoBd}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " & ClsPrefijo_FactStr.SstrNombreCampoBd & " = '" &
                lstrPrefPrefacturas & "'"
        Dim ldtbFechaPref = ClsPanorama.FdtbDataTable(ClsFactura.SstrNombreTabla, lstrCamposSelect,
                {{ClsFechaFacturaDtm.SstrNombreCampoBd, "DESC"}}, lstrFiltro)
        If ldtbFechaPref.Rows.Count > 0 Then
            Dim ldrwFecha As DataRow = ldtbFechaPref.Select()(0)
            ldtmFecha = ClsPanorama.FobjValorCampo(ldrwFecha(ClsFechaFacturaDtm.SstrNombreCampoBd),
                    EnuTipoValor.EnuDate)
        End If
        Return ldtmFecha
    End Function
    Private Shared Sub SActualiceEstadosFactAuto()
        Dim ldtbAntPorApli = FdtbAnticiposSinAplicar()
        Dim ldblIdCliente As Double, lstrIdPredioAgru As String
        Dim lobjAnti As New ClsAnticipo(EnuModoInstanciaObjDef.EnuUnico)
        Dim ldtbIdFactsFecha As DataTable
        For Each ldrwAntPorApl As DataRow In ldtbAntPorApli.Rows
            ldblIdCliente = ClsPanorama.FobjValorCampo(ldrwAntPorApl(
                    ClsIdCliente_AntDbl.SstrNombreCampoBd), EnuTipoValor.EnuDouble)
            lstrIdPredioAgru = ClsPanorama.FobjValorCampo(ldrwAntPorApl(
                    ClsIdPredioAgrupador_AntStr.SstrNombreCampoBd), EnuTipoValor.EnuString)
            ldtbIdFactsFecha = FdtbIdFactFechaPredio(ldblIdCliente, lstrIdPredioAgru,
                DtmFechaFactAuto)
            If ldtbIdFactsFecha.Rows.Count > 0 Then
                SActualiceEstadoFacturas(ldrwAntPorApl, ldtbIdFactsFecha)
            End If
        Next
    End Sub
    Private Shared Sub SActualiceEstadoFacturas(adrwAnticipo As DataRow,
            adtbFacturasFecha As DataTable)
        Dim lstrPrefFac As String, lentIdFact As Integer
        Dim lstrSerAnt As String = ClsPanorama.FobjValorCampo(adrwAnticipo(
                ClsServicios_AntStr.SstrNombreCampoBd), EnuTipoValor.EnuString)
        Dim ldecCrs As Decimal = ClsPanorama.FobjValorCampo(adrwAnticipo(
                ClsCreditos_AntDec.SstrNombreCampoBd), EnuTipoValor.EnuDecimal)
        Dim ldecDbs As Decimal = ClsPanorama.FobjValorCampo(adrwAnticipo(
                ClsDebitos_AntDec.SstrNombreCampoBd), EnuTipoValor.EnuDecimal)
        Dim ldecAntPorApli As Decimal = ldecCrs - ldecDbs
        Dim lobjFact As New ClsFactura()
        For Each ldrwFact As DataRow In adtbFacturasFecha.Rows
            lstrPrefFac = ClsPanorama.FobjValorCampo(ldrwFact(ClsPrefijo_FactStr.SstrNombreCampoBd),
                    EnuTipoValor.EnuString)
            If lstrPrefFac <> GCSTRPREFPREFACTURA Then
                lentIdFact = ClsPanorama.FobjValorCampo(ldrwFact(ClsIdFacturaEnt.SstrNombreCampoBd),
                    EnuTipoValor.EnuInteger)
                lobjFact.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefFac, lentIdFact})
                SActualiceEstado(lobjFact, ldecAntPorApli, lstrSerAnt)
            End If
        Next
    End Sub
    Private Shared Sub SActualiceEstado(aobjFactura As ClsFactura, adecValorAntPorApl As Decimal,
            astrSerAnt As String)
        Dim ldblIdCliente As Double = aobjFactura.ObjIdCliente_FactDbl.ObjValorPro
        Dim lstrIdPreAgr As String = aobjFactura.ObjIdPredioAgrupador_FacStr.ObjValorPro
        Dim lstrSerFac As String
        Dim lblnFactPorSer As Boolean, lblnActualizar As Boolean
        lblnFactPorSer = If(aobjFactura.ObjPredioAgrFactura IsNot Nothing,
                aobjFactura.ObjPredioAgrFactura.ObjFacturarPorServicio_PreBln.ObjValorPro,
                aobjFactura.ObjClienteFactura.ObjFactPorServicio_CliBln.ObjValorPro)
        Dim lobjItemFac As ClsItemFactura = aobjFactura.ColItemsFactura("1")
        If lblnFactPorSer Then
            lstrSerFac = If(lobjItemFac.ObjIdAno_ServicioItemFactShr.ObjValorPro > 0, "0",
                    lobjItemFac.ObjIdServicio_ItemFactShr.ToString())
            lblnActualizar = astrSerAnt.Contains(lstrSerFac)
        Else
            lblnActualizar = astrSerAnt.Contains("A")
            If Not lblnActualizar Then
                For Each lobjItemFac In aobjFactura.ColItemsFactura
                    If lobjItemFac.ObjIdAno_ServicioItemFactShr.ObjValorPro > 0 Then
                        lstrSerFac = "0"
                    Else
                        lstrSerFac = lobjItemFac.ObjIdServicio_ItemFactShr.ToString
                    End If
                    lblnActualizar = astrSerAnt.Contains(lstrSerFac)
                    If lblnActualizar Then Exit For
                Next
            End If
        End If
        If lblnActualizar Then
            Dim lobjEstado As ClsEstadoCuenta = aobjFactura.ObjEstadoCuenta
            If lobjEstado.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
                lobjEstado.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
            End If
            lobjEstado.ObjAntPorAplDec.ObjValorPro += adecValorAntPorApl
            lobjEstado.SActualice(True)
        End If
    End Sub
    ''' <summary>
    ''' Indica si en la fecha de la última generación de estado de cuenta se generaron estados
    ''' de cuenta sin factura asociada
    ''' </summary>
    ''' <returns></returns>    
#Region "DataTables utilizadas Estados de Cuenta"
    Private Shared Function FdtbIdFactFechaPredio(adblIdCliente As Double, astrIdPredAgru As String,
            adtmFechFact As Date) As DataTable
        Dim lstrTabla = ClsFactura.SstrNombreTabla
        Dim lstrCamSel = {ClsPrefijo_FactStr.SstrNombreCampoBd, ClsIdFacturaEnt.SstrNombreCampoBd}
        Dim lstrOrden = {{ClsPrefijo_FactStr.SstrNombreCampoBd, "ASC"},
                {ClsIdFacturaEnt.SstrNombreCampoBd, "ASC"}}
        Dim lstrFecha = "'" & ClsPanoramaDat.FstrFechaNormalizada(adtmFechFact) & "'"
        Dim lstrFiltro = StrFiltroUbicacion & " AND " & ClsIdCliente_FactDbl.SstrNombreCampoBd &
                " = " & adblIdCliente & " AND " & ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd &
                " = '" & astrIdPredAgru & "' AND " & ClsFechaFacturaDtm.SstrNombreCampoBd &
                " = " & lstrFecha & " AND " & ClsPrefijo_FactStr.SstrNombreCampoBd & " <> '" &
                GCSTRPREFPREFACTURA & "'"
        Dim ldtbIdFacsFechaPred = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamSel, lstrOrden, lstrFiltro)
        Return ldtbIdFacsFechaPred
    End Function
    Private Shared Function FdtbItemsFacVivas(adblIdCliente As Double,
            astrIdPreAgr As String, astrIdSer As String) As DataTable
        Dim lstrTablaPri = ClsItemFactura.SstrNombreTabla
        Dim lstrTablaSec = ClsFactura.SstrNombreTabla
        Dim lblnTodos = astrIdSer = "A", lshrIdSer As Short, lblnAdmin As Boolean
        If Not lblnTodos Then
            lshrIdSer = CShort(astrIdSer)
            lblnAdmin = lshrIdSer = 0
        End If
        Dim lstrCampSelePri = {"DISTINCT " & ClsPrefijo_ItemFactStr.SstrNombreCampoBd,
ClsIdFactura_ItemFactEnt.SstrNombreCampoBd,
ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd,
ClsIdServicio_ItemFactShr.SstrNombreCampoBd, "'F' as Proc"}
        Dim lstrCampSeleSec As String() = {ClsFechaFacturaDtm.SstrNombreCampoBd}
        Dim lstrCamRelPri As String() = {StrCampoCarpeta,
StrCampoCentroUtil, ClsPrefijo_ItemFactStr.SstrNombreCampoBd,
ClsIdFactura_ItemFactEnt.SstrNombreCampoBd}
        Dim lstrCamRelSec As String() = {StrCampoCarpeta,
StrCampoCentroUtil, ClsPrefijo_FactStr.SstrNombreCampoBd,
ClsIdFacturaEnt.SstrNombreCampoBd}
        Dim lstrFiltro = StrFiltroUbicacion_Pri & " AND " &
                ClsIdCliente_FactDbl.SstrNombreCampoBd & " = " & adblIdCliente &
                " AND " & ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd & " = '" &
                astrIdPreAgr & "' AND P." & ClsPrefijo_ItemFactStr.SstrNombreCampoBd &
                " <> '" & GCSTRPREFPREFACTURA & "' AND P." &
                ClsAnuladoBln.SstrNombreCampoBd & " = FALSE " & " AND P." &
                ClsDebitos_ItemFactDec.SstrNombreCampoBd & " > P." &
                ClsCreditos_ItemFactDec.SstrNombreCampoBd
        If Not lblnTodos Then
            If lblnAdmin Then
                lstrFiltro &= " AND " & ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd & " > 0"
            Else
                lstrFiltro &= " AND " & ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd & " = 0 AND " &
    ClsIdServicio_ItemFactShr.SstrNombreCampoBd & " = " & lshrIdSer
            End If
        End If
        Dim lstrOrden = {{ClsPrefijo_ItemFactStr.SstrNombreCampoBd, "ASC"},
{ClsIdFactura_ItemFactEnt.SstrNombreCampoBd, "ASC"}}
        Dim ldtbItemsFactViv = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCampSelePri, lstrTablaSec,
                lstrCampSeleSec, lstrCamRelPri, lstrCamRelSec, lstrOrden, False, lstrFiltro,
            Array.Empty(Of String))
        Return ldtbItemsFactViv
    End Function
    ' Si ashrIdAno = 9999 no filtra por año ni por servicio. Si ashrIdAno
    ' es igaul a cero solo tiene en cuenta los servicios permanente. Si ashrIdSer = 0 tiene en
    ' cuenta todos los servicios permanentes de lo contrario solo el servicio permanente indicado
    Private Shared Function FdtbItemsFacVivasPreAgr(astrIdPredAgr As String, ashridano As Short,
            ashrIdSer As Short)
        Dim lstrTablaPri = ClsItemFactura.SstrNombreTabla
        Dim lstrTablaSec = ClsFactura.SstrNombreTabla
        Dim lstrCampSelePri As String() = {"DISTINCT " & ClsPrefijo_ItemFactStr.SstrNombreCampoBd,
                ClsIdFactura_ItemFactEnt.SstrNombreCampoBd,
                ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd,
                ClsIdServicio_ItemFactShr.SstrNombreCampoBd, "'F' as Proc"}
        Dim lstrCampSeleSec As String() = {ClsIdCliente_FactDbl.SstrNombreCampoBd}
        Dim lstrCamRelPri As String() = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsPrefijo_ItemFactStr.SstrNombreCampoBd, ClsIdFactura_ItemFactEnt.SstrNombreCampoBd}
        Dim lstrCamRelSec As String() = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsPrefijo_FactStr.SstrNombreCampoBd, ClsIdFacturaEnt.SstrNombreCampoBd}
        Dim lstrFiltro = StrFiltroUbicacion_Pri & " AND " &
                ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd & " = '" & astrIdPredAgr &
                "' AND P." & ClsPrefijo_ItemFactStr.SstrNombreCampoBd & " <> '" &
                GCSTRPREFPREFACTURA & "' AND P." & ClsAnuladoBln.SstrNombreCampoBd &
                " = FALSE " & " AND P." & ClsDebitos_ItemFactDec.SstrNombreCampoBd &
                " > P." & ClsCreditos_ItemFactDec.SstrNombreCampoBd
        If ashridano = 9999 Then
            lstrFiltro = lstrFiltro
        ElseIf ashridano = 0 Then
            lstrFiltro &= " AND " & ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd & " = 0 "
            If ashrIdSer > 0 Then
                lstrFiltro &= " AND " & ClsIdServicio_ItemFactShr.SstrNombreCampoBd & " = " &
                        ashrIdSer
            End If
        End If
        Dim lstrOrden = {{ClsIdFactura_ItemFactEnt.SstrNombreCampoBd, "ASC"},
            {ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd, "DESC"},
            {ClsIdServicio_ItemFactShr.SstrNombreCampoBd, "ASC"},
            {ClsPrefijo_ItemFactStr.SstrNombreCampoBd, "ASC"},
            {ClsIdFactura_ItemFactEnt.SstrNombreCampoBd, "ASC"}}
        Dim ldtbItemsFactViv = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCampSelePri, lstrTablaSec,
                lstrCampSeleSec, lstrCamRelPri, lstrCamRelSec, lstrOrden, False, lstrFiltro,
                Array.Empty(Of String))
        Return ldtbItemsFactViv
    End Function
    Private Shared Function FdtbItemsFacVivas(adblIdCliente As Double, ashrIdSer As Short) As DataTable
        Dim lstrTablaPri = ClsItemFactura.SstrNombreTabla
        Dim lstrTablaSec = ClsFactura.SstrNombreTabla
        Dim lstrCampSelePri = {"DISTINCT " & ClsPrefijo_ItemFactStr.SstrNombreCampoBd,
                ClsIdFactura_ItemFactEnt.SstrNombreCampoBd,
                ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd,
                ClsIdServicio_ItemFactShr.SstrNombreCampoBd, "'F' as Proc"}
        Dim lstrCampSeleSec As String() = {ClsFechaFacturaDtm.SstrNombreCampoBd}
        Dim lstrCamRelPri As String() = {StrCampoCarpeta,
                StrCampoCentroUtil, ClsPrefijo_ItemFactStr.SstrNombreCampoBd,
                ClsIdFactura_ItemFactEnt.SstrNombreCampoBd}
        Dim lstrCamRelSec As String() = {StrCampoCarpeta,
                StrCampoCentroUtil, ClsPrefijo_FactStr.SstrNombreCampoBd,
                ClsIdFacturaEnt.SstrNombreCampoBd}
        Dim lstrFiltro = StrFiltroUbicacion_Pri & " AND " &
                ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd & " = '' AND " &
                ClsIdCliente_FactDbl.SstrNombreCampoBd & " = " & adblIdCliente &
                " AND P." & ClsPrefijo_ItemFactStr.SstrNombreCampoBd & " <> '" &
                GCSTRPREFPREFACTURA & "' AND P." & ClsAnuladoBln.SstrNombreCampoBd &
                " = FALSE " & " AND P." & ClsDebitos_ItemFactDec.SstrNombreCampoBd &
                " > P." & ClsCreditos_ItemFactDec.SstrNombreCampoBd & " AND " &
                ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd & " = 0 "
        If ashrIdSer > 0 Then
            lstrFiltro &= " AND " & ClsIdServicio_ItemFactShr.SstrNombreCampoBd & " = " & ashrIdSer
        End If
        Dim lstrOrden = {{ClsIdFactura_ItemFactEnt.SstrNombreCampoBd, "ASC"},
                {ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd, "DESC"},
                {ClsIdServicio_ItemFactShr.SstrNombreCampoBd, "ASC"},
                {ClsPrefijo_ItemFactStr.SstrNombreCampoBd, "ASC"},
                {ClsIdFactura_ItemFactEnt.SstrNombreCampoBd, "ASC"}}
        Dim ldtbItemsFactViv = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCampSelePri, lstrTablaSec,
                lstrCampSeleSec, lstrCamRelPri, lstrCamRelSec, lstrOrden, False, lstrFiltro,
                Array.Empty(Of String))
        Return ldtbItemsFactViv
    End Function
    Private Shared Function FdtbIdPrefactsClientes(adblIdCliente As Double) As DataTable
        Dim lstrTabla = ClsFactura.SstrNombreTabla
        Dim lstrCampSele As String() = {ClsIdFacturaEnt.SstrNombreCampoBd}
        Dim lstrOrden As String(,) = {{ClsIdFacturaEnt.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " & ClsPrefijo_FactStr.SstrNombreCampoBd & " = '" &
                GCSTRPREFPREFACTURA & "' AND " & ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd &
                " = '' AND " & ClsIdCliente_FactDbl.SstrNombreCampoBd & " = " & adblIdCliente
        Dim ldtbIdPreFacts = ClsPanorama.FdtbDataTable(lstrTabla, lstrCampSele, lstrOrden, lstrFiltro)
        Return ldtbIdPreFacts
    End Function
    Private Shared Function FdtbIdPrefactsPredAgr(astrIdPredAgr As String) As DataTable
        Dim lstrTabla = ClsFactura.SstrNombreTabla
        Dim lstrCampSele As String() = {ClsIdFacturaEnt.SstrNombreCampoBd}
        Dim lstrOrden As String(,) = {{ClsIdFacturaEnt.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " & ClsPrefijo_FactStr.SstrNombreCampoBd & " = '" &
                GCSTRPREFPREFACTURA & "' AND " & ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd &
                " = '" & astrIdPredAgr & "'"
        Dim ldtbIdPreFacts = ClsPanorama.FdtbDataTable(lstrTabla, lstrCampSele, lstrOrden, lstrFiltro)
        Return ldtbIdPreFacts
    End Function
#End Region
#End Region
#Region "Calculo y causación de intereses de Mora  a todas las deudas"
    ''' <summary>
    ''' Causa mora a todas las deudas el primer dia del período despues de cerrar mes o el día 
    ''' de hoy cuando está establecido que los documentos se deben generar con al fecha de hoy
    ''' </summary>
    ''' <param name="astrMens">Mensaje producido en el proceso</param>
    ''' <returns></returns>
    ' Causa mora en proceso FM
    Friend Function FblnCausoMoraGeneral(ByRef astrMens As String) As Boolean ' FM-
        Dim lblnNoHayError = False, lblnCausoMora As Boolean
        GobjPanDat.SControleProcesoObj(True)
        GblnCausandoFM = True
        Try
            Dim ldtmFechaCausacion = FdtmFechaCausaMoraGeneral()
            If ldtmFechaCausacion > GobjParametros.ObjFechaUltCausacionGralDtm.ObjValorPro Then
                If GobjParametros.FdblTasaMoraFecha(ldtmFechaCausacion) > 0 Then
                    GobjPanDat.SInicialiceTransaccion()
                    ' Clientes con predios agrupadores con deuda
                    Dim ldtbIdClientesConDeuda As DataTable = FdtbClientesConDeuda(False)
                    ObjArgumentoEventoPan.BlnCancele = False
                    ObjArgumentoEventoPan.DblCantProcesada = 0.0
                    ObjArgumentoEventoPan.DblCantAProcesar = ldtbIdClientesConDeuda.Rows.Count
                    ObjArgumentoEventoPan.EnuProceso = EnuProcesoDef.EnuCausaMora
                    RaiseEvent EvnInicio(Me, ObjArgumentoEventoPan)
                    Dim lobjCliente As ClsCliente = Nothing
                    Dim lobjValorLlave() As Object = Nothing
                    Dim ldblIdCliente = 0.0
                    If ldtbIdClientesConDeuda.Rows.Count > 0 Then
                        lobjCliente = New ClsCliente(EnuModoInstanciaObjDef.EnuUnico)
                    End If
                    For Each ldrwIdCli As DataRow In ldtbIdClientesConDeuda.Rows
                        ldblIdCliente = ClsPanorama.FobjValorCampo(ldrwIdCli(0),
                            EnuTipoValor.EnuDouble)
                        lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, ldblIdCliente}
                        lobjCliente.SAbra(lobjValorLlave)
                        If lobjCliente.BlnExiste Then
                            lobjCliente.SCauseMora(ldtmFechaCausacion, astrMens)
                            ObjArgumentoEventoPan.DblCantProcesada += 1
                            RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
                            If ObjArgumentoEventoPan.BlnCancele Then
                                Exit For
                            End If
                        Else
                            Throw New ErrorInesperadoPanLException("Cliente con Facturas vivas no existe!")
                        End If
                    Next
                End If
            End If
            If Not ObjArgumentoEventoPan.BlnCancele Then
                GobjParametros.SRegistreFechaUltCausa(ldtmFechaCausacion)
                GobjPanorama.SRegistreAccionLogApp("clsOrionCop", "Causacion General de Intereses")
            End If
            lblnNoHayError = True
        Catch ex As ErrorInesperadoPanDatException
            Throw
        Catch ex As ErrorInesperadoPanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As PanLException
            Throw
        Catch ex As ArgumentException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If Not lblnNoHayError OrElse ObjArgumentoEventoPan.BlnCancele Then
                lblnCausoMora = False
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            Else
                lblnCausoMora = True
                GobjPanDat.SConfirmeTransaccion()
                GobjPanDat.SControleProcesoObj(False)
            End If
            GblnCausandoFM = False
        End Try
        Return lblnCausoMora
    End Function
    Private Sub SCauseMoraPrefacturas(ByRef astrMens As String)
        Dim ldtmFechaFact As Date, i = 0, ldblIdCliente As Double, ldblCantAProcesar As Double
        Dim lentIdFac As Integer
        Dim lobjCliente As New ClsCliente(EnuModoInstanciaObjDef.EnuUnico)
        Dim lentIdFacMax = 0
        Dim ldtbPrefacturas = FdtbPrefacturas()
        If GobjParametros.BlnEFacAutorizado Then
            lentIdFacMax = GobjParametros.ObjRangoFraFinEnt.ObjValorPro
        End If
        ldblCantAProcesar = FdblCantidadPreFacAProcesar(ldtbPrefacturas)
        ObjArgumentoEventoPan.EnuProceso = EnuProcesoDef.EnuCausaMora
        ObjArgumentoEventoPan.DblCantAProcesar = ldblCantAProcesar
        RaiseEvent EvnInicio(Me, ObjArgumentoEventoPan)
        For Each ldrwFact As DataRow In ldtbPrefacturas.Rows
            i += 1
            ldblIdCliente = ClsPanorama.FobjValorCampo(
                    ldrwFact(ClsIdCliente_FactDbl.SstrNombreCampoBd), EnuTipoValor.EnuDouble)
            ldtmFechaFact = ClsPanorama.FobjValorCampo(ldrwFact(ClsFechaFacturaDtm.SstrNombreCampoBd),
                    EnuTipoValor.EnuDate)
            lentIdFac = ClsPanorama.FobjValorCampo(ldrwFact(ClsIdFacturaEnt.SstrNombreCampoBd),
                    EnuTipoValor.EnuInteger)
            If lentIdFacMax = 0 OrElse lentIdFac <= lentIdFacMax Then
                lobjCliente.SAbra({GshrIdCarpeta, GshrIdCentroUtil, ldblIdCliente})
                lobjCliente.SCauseMora(ldtmFechaFact, astrMens)
            End If
            ObjArgumentoEventoPan.DblCantProcesada = i
            RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
            If ObjArgumentoEventoPan.BlnCancele Then
                Exit For
            End If
        Next
    End Sub
    Friend Shared Function FblnHacerCierreMes() As Boolean
        Dim lblnCerrarMes = False
        If Not FblnHayItemsPorFacturar() Then
            Dim lstrPeriodoHoy = FstrPeriodoDeFecha(Date.Today)
            Dim lstrPeriodoActual = GobjParametros.ObjAnoActual.StrIdPeriodoActual
            Dim lblnPerActualFacturado =
                    GobjParametros.ObjAnoActual.ObjPeriodoActual.BlnPeriodoFacturado
            If Not GobjParametros.ObjAnoActual.ObjEstaCerradoAnoBln.ObjValorPro Then
                lblnCerrarMes = lblnPerActualFacturado AndAlso
                        lstrPeriodoHoy > lstrPeriodoActual
            End If
        End If
        Return lblnCerrarMes
    End Function
    Friend Shared Function FblnDebeCausarInt() As Boolean
        Dim lblnDebe = False
        Dim ldtmFechaUltCausaMora As Date = GobjParametros.ObjFechaUltCausacionGralDtm.ObjValorPro
        If GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro Then
            If Not GobjParametros.ObjAnoActual.FblnFacturacionGenerada Then
                lblnDebe = ldtmFechaUltCausaMora < Date.Today
            End If
        Else
            lblnDebe = FblnHacerCierreMes()
        End If
        Return lblnDebe
    End Function
    Friend Shared Function FdtmFechaCausaMoraGeneral() As Date
        Dim ldtmFechaCausa As Date
        If GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro Then
            ldtmFechaCausa = Date.Today
        Else
            ldtmFechaCausa = GobjParametros.ObjAnoActual.ObjPeriodoActual.
                    DtmFechaInicioPeriodo
        End If
        Return ldtmFechaCausa
    End Function
    Friend Shared Function FblnDebeCrearAno() As Boolean
        Dim lshrIdPeriodoHoy = Today.Month, lblnDebe As Boolean
        Dim lshrIdAnoHoy = Today.Year
        Dim lshrIdPeriodoActual =
                GobjParametros.ObjAnoActual.ObjPeriodoActual.ObjIdPeriodoShr.ObjValorPro
        lblnDebe = lshrIdPeriodoHoy >= 1 AndAlso lshrIdPeriodoActual = 12 AndAlso
                Not (GobjParametros.ColAnos.Contains(lshrIdAnoHoy.ToString))
        Return lblnDebe
    End Function
#End Region
#Region "Importar Notas de Intereses"
    Friend Shared Function FblnPuedeImpNsDb(ByRef astrMens As String) As Boolean
        Try
            Dim lblnPuede = True
            Dim lstrArchivo = GstrTrayDatPrg & "PlantillaNotasDb_OrionPLus.xlsx"
            lblnPuede = My.Computer.FileSystem.FileExists(lstrArchivo)
            If Not lblnPuede Then
                astrMens = "No existe el Archivo con los Datos a importar en la Ubicación esperada!"
                Return lblnPuede
            End If
            Dim ldtbItemsNsDb = ClsPanorama.FdtbTablaAccess(lstrArchivo, "", "Plantilla$")
            lblnPuede = ldtbItemsNsDb IsNot Nothing
            If Not lblnPuede Then
                astrMens = "El Archivo no contiene la hoja con los datos a importar!"
                Return lblnPuede
            End If
            lblnPuede = ldtbItemsNsDb.Rows.Count > 0
            If Not lblnPuede Then
                astrMens = "La hoja de Excel esta vacia!"
                Return lblnPuede
            End If
            lblnPuede = Not ClsCentroUtilOriCop.FblnHayNotasDb
            If Not lblnPuede Then
                lblnPuede = FblnEsPrimeraDelMes(ldtbItemsNsDb)
            End If
            If Not lblnPuede Then
                astrMens = "Ya hay notas de intereses generadas o no es la primera causación del mes!"
                Return lblnPuede
            End If
            Return lblnPuede
        Catch ex As PanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' Indica si la causación de intereses general a llevarse a cabo sera la primera del mes
    ''' para lo cual la fecha de la últrima causación debe  pertenecer al periodo actual
    ''' </summary>
    ''' <returns></returns>
    Friend Shared Function FblnEsPrimeraDelMes(adtbItemsNsDb As DataTable) As Boolean
        Dim lblnEsPrimera As Boolean
        Dim ldtmFechaNotas As Date = ClsPanorama.FobjValorCampo(adtbItemsNsDb(0)("FechaCauso"),
                    EnuTipoValor.EnuDate)
        Dim lblnCierreMes As Boolean = ClsPanorama.FobjValorCampo(adtbItemsNsDb(0)("CierreMes"),
                    EnuTipoValor.EnuBoolean)
        Dim ldtmFechaUltCausa As Date = GobjParametros.ObjFechaUltCausacionGralDtm.ObjValorPro
        Dim lstrPeriodoActual = GobjParametros.ObjAnoActual.StrIdPeriodoActual
        Dim lstrPeriodoHoy = FstrPeriodoDeFecha(Today)
        Dim lstrPeriodoUltCausa = FstrPeriodoDeFecha(ldtmFechaUltCausa)
        If GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro Then
            lblnEsPrimera = lstrPeriodoActual = lstrPeriodoHoy AndAlso lstrPeriodoUltCausa <
                        lstrPeriodoActual AndAlso Not lblnCierreMes
        Else
            lblnEsPrimera = lstrPeriodoActual = lstrPeriodoUltCausa AndAlso lblnCierreMes
        End If
        Return lblnEsPrimera
    End Function
    Friend Shared Function FblnImportoNotasDb() As Boolean
        Dim lblnNoHayError = False, lblnImportoNDb = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            BlnProcesoEspecial = True
            GobjPanDat.SInicialiceTransaccion()
            SImporteNotasDb()
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
            BlnProcesoEspecial = False
            If lblnNoHayError Then
                GobjPanDat.SConfirmeTransaccion()
                GobjPanDat.SControleProcesoObj(False)
                lblnImportoNDb = True
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
        Return lblnImportoNDb
    End Function
    Private Shared Sub SImporteNotasDb()
        Dim lobjCliente = New ClsCliente(EnuModoInstanciaObjDef.EnuUnico)
        Dim lobjFact = New ClsFactura()
        Dim lobjItemFac As ClsItemFactura
        Dim ldblIdCliente As Double, lstrPreAgr As String
        Dim lstrPrefFac As String, lentIdFac As Integer, lentIdItemFac As Integer
        Dim lentIdFacCauso = 0, lentDiasMora As Integer
        Dim ldecBase As Decimal, ldblTasa As Double, ldecVlrMora As Decimal, ldecVlrMoraFac = 0D
        Dim lobjNDb As ClsNotaDb = Nothing
        Dim lobjLlave As Object()
        Dim lstrArchivo = GstrTrayDatPrg & "PlantillaNotasDb_OrionPLus.xlsx"
        Dim ldtbItemsNsDb = ClsPanorama.FdtbTablaAccess(lstrArchivo, "", "Plantilla$")
        Dim ldtmFechaCauso As Date = ClsPanorama.FobjValorCampo(ldtbItemsNsDb(0)("FechaCauso"),
                EnuTipoValor.EnuDate)
        Dim lblnCierraMes = ClsPanorama.FobjValorCampo(ldtbItemsNsDb(0)("CierreMes"),
                EnuTipoValor.EnuBoolean)
        For Each ldrwItemNdb As DataRow In ldtbItemsNsDb.Rows
            ldblIdCliente = ClsPanorama.FobjValorCampo(ldrwItemNdb("IdCliente"),
                    EnuTipoValor.EnuDouble)
            lstrPreAgr = ClsPanorama.FobjValorCampo(ldrwItemNdb("IdPreAgr"),
                    EnuTipoValor.EnuString)
            lstrPrefFac = ClsPanorama.FobjValorCampo(ldrwItemNdb("PrefFactura"),
                    EnuTipoValor.EnuString)
            lentIdFac = ClsPanorama.FobjValorCampo(ldrwItemNdb("IdFactura"),
                    EnuTipoValor.EnuInteger)
            lentIdItemFac = ClsPanorama.FobjValorCampo(ldrwItemNdb("IdItemFac"),
                    EnuTipoValor.EnuInteger)
            ldecBase = ClsPanorama.FobjValorCampo(ldrwItemNdb("Base"),
                    EnuTipoValor.EnuDecimal)
            ldblTasa = ClsPanorama.FobjValorCampo(ldrwItemNdb("Tasa"),
                    EnuTipoValor.EnuDouble)
            ldecVlrMora = ClsPanorama.FobjValorCampo(ldrwItemNdb("Valor"),
                    EnuTipoValor.EnuDecimal)
            lentDiasMora = ClsPanorama.FobjValorCampo(ldrwItemNdb("DiasMora"),
                    EnuTipoValor.EnuInteger)
            If IsNothing(lstrPrefFac) Then lstrPrefFac = String.Empty
            If FblnPuedeCausar(ldrwItemNdb) Then
                If lentIdFacCauso <> lentIdFac Then
                    If lentIdFacCauso > 0 Then
                        If ldecVlrMoraFac > 0 Then
                            SActuaDocsNdb(lobjFact, lobjNDb, ldecVlrMoraFac)
                            ldecVlrMoraFac = 0
                        End If
                    End If
                    If ldecVlrMora > 0 Then
                        lobjLlave = {GshrIdCarpeta, GshrIdCentroUtil, ldblIdCliente}
                        lobjCliente.SAbra(lobjLlave)
                        lobjLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrPrefFac, lentIdFac}
                        lobjFact.SAbra(lobjLlave)
                        lentIdFacCauso = lentIdFac
                        lobjNDb = lobjCliente.FobjNuevaNotaDb(ldtmFechaCauso, lstrPreAgr,
                                EnuOrigenNotaDb.EnuImportado)
                    End If
                End If
                If ldecVlrMora > 0 Then
                    lobjItemFac = lobjFact.ColItemsFactura(lentIdItemFac.ToString)
                    If lobjItemFac.ObjServicio.FblnCausaMora Then
                        lobjNDb.SAdicioneItemNotaDb(lentIdFac, lstrPrefFac, ldecBase, lentIdItemFac,
                        ldblTasa, ldecVlrMora, lentDiasMora, ldtmFechaCauso)
                        lobjItemFac.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
                        lobjItemFac.ObjFechaCausoIntMora_Dtm.ObjValorPro = ldtmFechaCauso
                        lobjItemFac.ObjDebitos_ItemFactDec.ObjValorPro += ldecVlrMora
                        ldecVlrMoraFac += ldecVlrMora
                    Else
                        ldecVlrMora = 0
                    End If
                End If
            Else
                If ldecVlrMoraFac > 0 Then
                    SActuaDocsNdb(lobjFact, lobjNDb, ldecVlrMoraFac)
                    ldecVlrMoraFac = 0
                End If
            End If
        Next
        If ldecVlrMoraFac > 0 Then
            SActuaDocsNdb(lobjFact, lobjNDb, ldecVlrMoraFac)
        End If
        GobjParametros.SRegistreFechaUltCausa(ldtmFechaCauso)
        GobjPanorama.SRegistreAccionLogApp("clsOrionCop", "Importación General de Intereses")
        If lblnCierraMes Then
            SActuFechasCausoMora(ldtmFechaCauso)
            GobjParametros.SCierrePeriodo()
        End If
    End Sub
    Private Shared Sub SActuFechasCausoMora(adtmFechaCierre As Date)
        Dim lstrFecha = "'" & ClsPanoramaDat.FstrFechaNormalizada(adtmFechaCierre.ToString) & "'"
        Dim lstrTabla = ClsItemFactura.SstrNombreTabla
        Dim lstbSql As New StringBuilder
        With lstbSql
            .Clear()
            .Append("UPDATE ").Append(lstrTabla).Append(" SET ")
            .Append(ClsFechaCausoIntMora_Dtm.SstrNombreCampoBd).Append(" = ").Append(lstrFecha)
            .Append(" WHERE ").Append(StrCampoCarpeta).Append(" = ")
            .Append(GshrIdCarpeta).Append(" AND ").Append(StrCampoCentroUtil)
            .Append(" = ").Append(GshrIdCentroUtil).Append(" AND ")
            .Append(ClsDebitos_ItemFactDec.SstrNombreCampoBd).Append(" > ")
            .Append(ClsCreditos_ItemFactDec.SstrNombreCampoBd)
        End With
        Dim lstrSql = lstbSql.ToString
        GobjPanDat.SEjecuteSentenciaSql(lstrSql)
    End Sub
    Private Shared Sub SActuaDocsNdb(aobjFactura As ClsFactura, aobjNotaDb As ClsNotaDb,
            adecVlrMoraFac As Decimal)
        aobjFactura.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
        aobjFactura.ObjDebitos_FactDec.ObjValorPro += adecVlrMoraFac
        aobjFactura.SActualice(True)
        If GobjParametros.BlnEFacAutorizado Then
            If aobjFactura.FenuVerFacEFac = EnuVerEFac.EnuV1 Then
                aobjNotaDb.ObjVerEFacEnt.ObjValorPro = EnuVerEFac.EnuV1
            Else
                aobjNotaDb.ObjVerEFacEnt.ObjValorPro = EnuVerEFac.EnuV2
            End If
        End If
        aobjNotaDb.SActualice(True)
    End Sub
    ' Verifica que el estado de la cuenta no está en "Perdída" 
    Private Shared Function FblnPuedeCausar(adrwItemNDb As DataRow) As Boolean
        Dim ldblIdCliente As Double = ClsPanorama.FobjValorCampo(adrwItemNDb("IdCliente"),
                EnuTipoValor.EnuDouble)
        Dim lstrPreAgr As String = ClsPanorama.FobjValorCampo(adrwItemNDb("IdPreAgr"),
                EnuTipoValor.EnuString)
        Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, ldblIdCliente}
        Dim lobjCliente As New ClsCliente(EnuModoInstanciaObjDef.EnuUnico)
        lobjCliente.SAbra(lobjValorLlave)
        Dim lblnPuede = lobjCliente.FblnPuedeCausarMora(lstrPreAgr)
        Return lblnPuede
    End Function
#End Region
#Region "Ajuste Cuotas Administración"
    Friend Sub SAjusteCuotas(aobjAno As ClsAno, aentCantidadCuotas As Integer,
            ablnNoAjustar As Boolean)
        Dim lblnCancelo = False, lblnNoHayError = False
        Dim lobjServicioAjuste As ClsServicio = Nothing
        Try
            GobjPanDat.SControleProcesoObj(True)
            GobjPanDat.SInicialiceTransaccion()
            ' Examino si hay Servicios de Ajuste generados ne el periodo actual
            ' para suprimirlos
            For Each lobjServicio As ClsServicio In aobjAno.ColServiciosAno
                If lobjServicio.ObjEsAjusteBln.ObjValorPro Then
                    If lobjServicio.ObjPeriodoInicioStr.ObjValorPro =
                            aobjAno.StrIdPeriodoActual Then
                        SSuprimaSerAjuste(lobjServicio)
                    End If
                End If
            Next
            ' Genero los nuevos servicios de ajuste con los items del programa de
            ' facturación y las notas de ajuste si hubiere lugar
            For Each lobjServicio As ClsServicio In aobjAno.ColServiciosAno
                If Not lobjServicio.ObjEsAjusteBln.ObjValorPro Then
                    If Not ablnNoAjustar Then
                        lobjServicioAjuste = aobjAno.FobjServicioAjuste(lobjServicio,
                                aentCantidadCuotas)
                        SGenereAjustesPredios(lobjServicio, lobjServicioAjuste,
                                aentCantidadCuotas, lblnCancelo)
                        If lblnCancelo Then Exit For
                    End If
                    If lobjServicio.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
                        With lobjServicio
                            .EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
                            .ObjEstaAjustadoBln.ObjValorPro = True
                            .SActualice(True)
                        End With
                    End If
                End If
            Next
            lblnNoHayError = True
        Catch ex As ErrorInesperadoPanDatException
            Throw
        Catch ex As ErrorInesperadoPanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As PanLException
            Throw
        Catch ex As ArgumentException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                If Not lblnCancelo Then
                    GobjPanDat.SConfirmeTransaccion()
                Else
                    GobjPanDat.SAborteTransaccion()
                End If
                GobjPanDat.SControleProcesoObj(False)
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    Private Sub SGenereAjustesPredios(aobjServicio As ClsServicio,
            aobjServicioAjuste As ClsServicio, aentCantidadCuotas As Integer,
            ByRef ablnCancelo As Boolean)
        Dim i = 0.0
        Dim lstrPeriodoFin As String = aobjServicioAjuste.ObjPeriodoInicioStr.ObjValorPro
        Dim ldtbValorFacturadoPorPredio = FdtbValorFacturadoPorPredio(aobjServicio,
                lstrPeriodoFin)
        Dim ldtbValorNotasAjuste = FdtbValorNotasAjuste()
        Dim ldtbPreConDscto = FdtbPrediosConDsctoCA()
        Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.EnuNavegable)
        Dim lobjModuloCon As ClsModuloContribucion
        lobjPredio.SVayaAlPrimero()
        ObjArgumentoEventoPan.BlnCancele = False
        ObjArgumentoEventoPan.DblCantAProcesar = FentCantidadPredios()
        ObjArgumentoEventoPan.EnuProceso = EnuProcesoDef.EnuAjusCuota
        RaiseEvent EvnInicio(Me, ObjArgumentoEventoPan)
        Dim ldrwPredio As DataRow, lstrSectorPredio As String, lblnCalcule As Boolean
        Dim lstrPredio As String, lstrFiltro As String, lstrFiltro1 As String, lstrFiltro2 As String
        Dim ldecValorFac As Decimal, ldecValorReint As Decimal, ldecValorDebio As Decimal
        Dim ldecValorAjuste As Decimal, ldecValorDscto As Decimal, lshrIdModulo As Short
        Do While lobjPredio.BlnExiste
            i += 1
            ldecValorFac = 0
            ldecValorReint = 0
            ldecValorDscto = 0
            lstrSectorPredio = lobjPredio.ObjSector.ObjIdSectorShr.ToString
            For Each lobjModuSer As ClsModuloServicio In aobjServicio.ColModulosServicio
                lshrIdModulo = lobjModuSer.ObjIdModulo_ModuloServicioShr.ObjValorPro
                lobjModuloCon = GobjParametros.ColModulos(lshrIdModulo.ToString())
                lblnCalcule = lobjModuloCon.ColSectoresModulo.Contains(lstrSectorPredio)
                If lblnCalcule Then Exit For
            Next
            If lblnCalcule Then
                lstrPredio = lobjPredio.ObjIdPredioStr.ObjValorPro
                lstrFiltro = ClsIdPredioStr.SstrNombreCampoBd & " = '" & lstrPredio & "'"
                lstrFiltro1 = ClsIdPredio_NotaAjusteStr.SstrNombreCampoBd & " = '" & lstrPredio & "'"
                lstrFiltro2 = ClsIdPredio_ItemFactStr.SstrNombreCampoBd & " = '" & lstrPredio & "'"
                If ldtbValorFacturadoPorPredio.Select(lstrFiltro).Length > 0 Then
                    ldrwPredio = ldtbValorFacturadoPorPredio.Select(lstrFiltro)(0)
                    ldecValorFac = ClsPanorama.FobjValorCampo(ldrwPredio(
                            ClsValor_ItemFactDec.SstrNombreCampoBd),
                            EnuTipoValor.EnuDecimal)
                    If GobjParametros.ColAnos.Count = 1 Then
                        Dim lstrPerFin = GobjParametros.ObjAnoActual.StrIdPeriodoActual
                        Dim lentCanPerFact = FentCanPeriodosFacturados(GobjParametros.ObjAnoActual)
                        Dim lentTotCantPer As Integer = CType(Right(lstrPerFin, 2), Integer) - 1
                        ldecValorFac = ldecValorFac * lentTotCantPer / lentCanPerFact
                    End If
                End If
                If ldtbValorNotasAjuste.Select(lstrFiltro1).Length > 0 Then
                    ldrwPredio = ldtbValorNotasAjuste.Select(lstrFiltro1)(0)
                    ldecValorReint = ClsPanorama.FobjValorCampo(ldrwPredio(
                            ClsValor_NotaAjusteDec.SstrNombreCampoBd),
                            EnuTipoValor.EnuDecimal)
                End If
                Dim ldrwPredios As DataRow() = ldtbPreConDscto.Select(lstrFiltro2)
                If ldrwPredios.Length > 0 Then
                    If GobjParametros.ObjAnoActual.ObjAplicaDsctoPPBln.ObjValorPro Then
                        ldecValorDscto = FdecSaldoDsctoPredio_NoPP(lstrPredio)
                    Else
                        ldecValorDscto = FdecSaldoDsctoPredio(lstrPredio)
                    End If
                End If
                ldecValorFac -= ldecValorReint - ldecValorDscto
                ldecValorDebio = FdecValorDebioFacturarAPredio(aobjServicio,
                        lstrPeriodoFin, lstrPredio)
                ldecValorAjuste = ldecValorDebio - ldecValorFac
                If ldecValorAjuste <> 0 Then
                    Dim ldecValorCuotaAjuste = FdecValorRedondeado(ldecValorAjuste /
                            aentCantidadCuotas)
                    If ldecValorCuotaAjuste > 0 Then
                        SCreeItemProgramaFactPredio(aobjServicioAjuste,
                                ldecValorCuotaAjuste, lobjPredio)
                    ElseIf ldecValorAjuste < 0 Then
                        ldecValorAjuste *= -1
                        Dim lstrIdCtaDb = aobjServicio.ObjCodigoCuentaCrStr.ObjValorPro
                        Dim lstrSer As String
                        If GobjParametros.ObjPermiteAnticipoPorServicioBln.ObjValorPro Then
                            lstrSer = If(aobjServicio.ObjIdTipoServicioByt.ObjValorPro =
                                    EnuTipoServicio.EnuAnual, "0",
                                    aobjServicio.ObjIdServicioShr.ToString)
                        Else
                            lstrSer = "A"
                        End If
                        SGenereNotaAjuste(lobjPredio, ldecValorAjuste, lstrSer, lstrIdCtaDb)
                    End If
                End If
            End If
            ObjArgumentoEventoPan.DblCantProcesada = i
            RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
            If ObjArgumentoEventoPan.BlnCancele Then
                ablnCancelo = True
                Exit Do
            End If
            lobjPredio.SVayaAlSiguiente()
        Loop
    End Sub
    Private Shared Sub SGenereNotaAjuste(aobjPredio As ClsPredio, adecValorAjuste As Decimal,
            astrServicio As String, astrIdCtaDb As String)
        Dim lstrIdPredioAgr As String = aobjPredio.ObjIdPredioAgrupadorStr.ObjValorPro
        Dim lobjPredioAgr As ClsPredio = aobjPredio.ObjPredioAgrupador
        Dim ldtmFechaNotaAjuste = GCDTMFECHANULA
        Dim ldecVlrAjuste As Decimal, ldecVlrAjustado = 0D, i = 0
        Dim lobjCliente As ClsCliente, ldblPorcPart As Double
        If GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro Then
            ldtmFechaNotaAjuste = Date.Today
        Else
            ldtmFechaNotaAjuste =
                    GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
        End If
        If aobjPredio.ObjIdTipoDestinatarioFacturaByt.ObjValorPro =
                EnuDestinatarioFacturaDef.EnuPropietario Then
            For Each lobjProp As ClsPropietario In aobjPredio.ColPropietarios
                i += 1
                lobjCliente = lobjProp.ObjCliente
                ldblPorcPart = lobjProp.ObjPorcentajePartiDbl.ObjValorPro
                ldecVlrAjuste = ClsOrionCop.FdecValorRedondeado(adecValorAjuste *
                        ldblPorcPart)
                If i = aobjPredio.ColPropietarios.Count Then
                    If ldecVlrAjustado <> adecValorAjuste Then
                        ldecVlrAjuste = adecValorAjuste - ldecVlrAjustado
                    End If
                End If
                ldecVlrAjustado += ldecVlrAjuste
                If ldecVlrAjuste > 0 Then
                    Dim lobjNotaAjuste = lobjCliente.FobjNuevaNotaAjuste()
                    With lobjNotaAjuste
                        .StrServicios &= astrServicio
                        .StrIdCuentaDb = astrIdCtaDb
                        .ObjFecha_NotaAjusteDtm.ObjValorPro = ldtmFechaNotaAjuste
                        .ObjIdPredio_NotaAjusteStr.ObjValorPro = aobjPredio.ObjIdPredioStr.ObjValorPro
                        .ObjIdPredioAgrupador_NotaAjusteStr.ObjValorPro = lstrIdPredioAgr
                        .ObjValor_NotaAjusteDec.ObjValorPro = ldecVlrAjuste
                        .SActualice(True)
                    End With
                End If
            Next
        Else
            If adecValorAjuste > 0 Then
                lobjCliente = aobjPredio.ObjArrendatario
                Dim lobjNotaAjuste = lobjCliente.FobjNuevaNotaAjuste()
                With lobjNotaAjuste
                    .StrServicios &= astrServicio
                    .StrIdCuentaDb = astrIdCtaDb
                    .ObjFecha_NotaAjusteDtm.ObjValorPro = ldtmFechaNotaAjuste
                    .ObjIdPredio_NotaAjusteStr.ObjValorPro = aobjPredio.ObjIdPredioStr.ObjValorPro
                    .ObjIdPredioAgrupador_NotaAjusteStr.ObjValorPro = lstrIdPredioAgr
                    .ObjValor_NotaAjusteDec.ObjValorPro = adecValorAjuste
                    .SActualice(True)
                End With
            End If
        End If
    End Sub
    Private Shared Function FdtbValorFacturadoPorPredio(aobjServicio As ClsServicio,
            astrIdPeriodoFin As String) As DataTable
        Dim lstbExpSql As New StringBuilder
        Dim lstrPeriodoIni = aobjServicio.ObjPeriodoInicioStr.ToString
        Dim lshrIdServicio As Short = aobjServicio.ObjIdServicioShr.ObjValorPro
        Dim lshrIdAno As Short = aobjServicio.ObjIdAno_ServicioShr.ObjValorPro
        Dim lstrTablaPri = ClsFactura.SstrNombreTabla
        Dim lstrTablaSec = ClsItemFactura.SstrNombreTabla
        Dim lstrCampSelePri As String() = {}
        Dim lstrCampSeleSec As String() = {ClsIdPredio_ItemFactStr.SstrNombreCampoBd,
                "SUM(" & ClsValor_ItemFactDec.SstrNombreCampoBd & ")"}
        Dim lstrCamRelPri As String() = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsPrefijo_FactStr.SstrNombreCampoBd, ClsIdFacturaEnt.SstrNombreCampoBd}
        Dim lstrCamRelSec As String() = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsPrefijo_ItemFactStr.SstrNombreCampoBd, ClsIdFactura_ItemFactEnt.SstrNombreCampoBd}
        ' Filtro
        Dim lstrFiltroSer = FstrFiltroSer(aobjServicio)
        With lstbExpSql
            .Append(StrFiltroUbicacion_Pri).Append(" AND ").Append(lstrFiltroSer).Append(" AND ")
            .Append(ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd).Append(" = ").Append(lshrIdAno)
            .Append(" AND ").Append(ClsPeriodo_ItemFactStr.SstrNombreCampoBd).Append(" >= '")
            .Append(lstrPeriodoIni).Append("'").Append(" AND P.")
            .Append(ClsAnuladoBln.SstrNombreCampoBd).Append(" = FALSE AND ")
            .Append(ClsPeriodo_ItemFactStr.SstrNombreCampoBd).Append(" < '")
            .Append(astrIdPeriodoFin).Append("' AND ").Append(ClsIdModoFacturacionByt.SstrNombreCampoBd)
            .Append(" <> ").Append(EnuModoFacturacionDef.EnuImportada)
        End With
        Dim lstrFiltro = lstbExpSql.ToString
        Dim lstrOrden(,) = {{ClsIdPredio_ItemFactStr.SstrNombreCampoBd, "ASC"}}
        Dim lstrGrupo As String() = {ClsIdPredio_ItemFactStr.SstrNombreCampoBd}
        Dim ldtbRes = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCampSelePri, lstrTablaSec,
                lstrCampSeleSec, lstrCamRelPri, lstrCamRelSec, lstrOrden, True, lstrFiltro, lstrGrupo)
        Return ldtbRes
    End Function
    Private Shared Function FstrFiltroSer(aobjServicio As ClsServicio) As String
        Dim lobjAno As ClsAno = aobjServicio.ObjMiAno
        Dim lshrIdSer As Short = aobjServicio.ObjIdServicioShr.ObjValorPro
        Dim lshrIdSerAju As Short
        Dim lstrFiltro = "(" & ClsIdServicio_ItemFactShr.SstrNombreCampoBd & " = " &
                lshrIdSer
        For Each lobjSer As ClsServicio In lobjAno.ColServiciosAno
            If lobjSer.ObjEsAjusteBln.ObjValorPro AndAlso
                    lobjSer.ObjIdServicioAjustadoShr.ObjValorPro = lshrIdSer Then
                lshrIdSerAju = lobjSer.ObjIdServicioShr.ObjValorPro
                lstrFiltro &= " OR " & ClsIdServicio_ItemFactShr.SstrNombreCampoBd &
                        " = " & lshrIdSerAju
            End If
        Next
        lstrFiltro &= ")"
        Return lstrFiltro
    End Function
    Private Shared Function FdtbPrediosConDsctoCA() As DataTable
        Dim lstrFechaIni As String = "'" & ClsPanoramaDat.FstrFechaNormalizada(
                GobjParametros.ObjAnoActual.DtmFechaInicioAno) & "'"
        Dim lstrFechaFin As String = "'" & ClsPanoramaDat.FstrFechaNormalizada(Date.Today) & "'"
        Dim lentIdAno = Date.Today.Year
        Dim lstrTblPri = ClsItemFactura.SstrNombreTabla
        Dim lstrTblSec = ClsNovedad.SstrNombreTabla
        Dim lstrCamSelPri = {"DISTINCT " & ClsIdPredio_ItemFactStr.SstrNombreCampoBd}
        Dim lstrCamSelSec As String() = {}
        Dim lstrCamposRelPri = {StrCampoCarpeta,
                StrCampoCentroUtil,
                ClsIdFactura_ItemFactEnt.SstrNombreCampoBd,
                ClsIdItemFacturaShr.SstrNombreCampoBd}
        Dim lstrCamposRelSec = {StrCampoCarpeta,
                StrCampoCentroUtil,
                ClsIdFactura_NovEnt.SstrNombreCampoBd,
                ClsIdItemFacturaShr.SstrNombreCampoBd}
        Dim lstrCamposGrup As String() = {}
        Dim lstrOrden = {{"", ""}}
        Dim lstrFiltro = StrFiltroUbicacion_Pri &
                " AND P." & ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd & " = " &
                lentIdAno.ToString & " AND " & ClsIdTipoNovedadByt.SstrNombreCampoBd &
                " = " & EnuTipoNov.EnuCrDctoCap & " AND " &
                ClsFechaNovedadDtm.SstrNombreCampoBd & " BETWEEN " & lstrFechaIni &
                " AND " & lstrFechaFin
        Dim ldtbPredConDscoCA = ClsPanorama.FdtbDataTable(lstrTblPri, lstrCamSelPri, lstrTblSec,
                lstrCamSelSec, lstrCamposRelPri, lstrCamposRelSec, lstrOrden, lstrFiltro, lstrCamposGrup, True)
        Return ldtbPredConDscoCA
    End Function
    ''' <summary>
    ''' Devuelve el valor descontado a los servicios de cuota de administración desde el primero del año
    ''' actual para tenerlo en cuenta en el calculo del ajuste cuando la nueva cuota es menor a lo que
    ''' se estaba conrando desde enros del año.Se utiliza solo cuando no hay descuentos por pronto pago!
    ''' </summary>
    ''' <param name="astrIdPredio">Id del Predio para el cual se calcula el valor del descuento</param>
    ''' <returns></returns>
    Private Shared Function FdecSaldoDsctoPredio(astrIdPredio As String) As Decimal
        Dim ldecDscto As Decimal = FdecDsctosPredio(astrIdPredio)
        Dim ldecRevDscto As Decimal = FdecRevDsctosPredio(astrIdPredio)
        Dim ldecSaldoDscto = ldecRevDscto - ldecRevDscto
        Return ldecDscto - ldecSaldoDscto
    End Function
    Private Shared Function FdecDsctosPredio(astrIdPredio As String) As Decimal
        Dim ldecDscto = 0D
        Dim lstrFechaIni As String = "'" & ClsPanoramaDat.FstrFechaNormalizada(
                GobjParametros.ObjAnoActual.DtmFechaInicioAno) & "'"
        Dim lstrFechaFin As String = "'" & ClsPanoramaDat.FstrFechaNormalizada(
                Date.Today) & "'"
        Dim lstrTablaPri = ClsNovedad.SstrNombreTabla
        Dim lstrTablaSec = ClsItemFactura.SstrNombreTabla
        Dim lstrCampSelPro As String() = {"SUM(" & ClsValor_NovDec.SstrNombreCampoBd & ")"}
        Dim lstrCampSelSec As String() = {}
        Dim lstrCampRelPri As String() = {StrCampoCarpeta, StrCampoCentroUtil,
                    ClsPrefijoFact_NovStr.SstrNombreCampoBd, ClsIdFactura_NovEnt.SstrNombreCampoBd,
                    ClsIdItemDocOrigen_NovShr.SstrNombreCampoBd}
        Dim lstrCampRelSec As String() = {StrCampoCarpeta, StrCampoCentroUtil,
                    ClsPrefijo_ItemFactStr.SstrNombreCampoBd,
                    ClsIdFactura_ItemFactEnt.SstrNombreCampoBd,
                    ClsIdItemFacturaShr.SstrNombreCampoBd}
        Dim lstrFiltro = StrFiltroUbicacion_Pri & " AND " & ClsIdPredio_ItemFactStr.SstrNombreCampoBd &
                " = '" & astrIdPredio & "' AND P." & ClsIdAno_NovShr.SstrNombreCampoBd & " = " &
                 Date.Today.Year & " AND " & ClsFechaNovedadDtm.SstrNombreCampoBd & " BETWEEN " &
                 lstrFechaIni & " AND " & lstrFechaFin & " AND " &
                 ClsIdTipoNovedadByt.SstrNombreCampoBd & " = " & EnuTipoNov.EnuCrDctoCap
        Dim ldtbRes = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCampSelPro, lstrTablaSec,
                lstrCampSelSec, lstrCampRelPri, lstrCampRelSec, {{"", ""}}, True, lstrFiltro, {})
        If ldtbRes.Rows.Count > 0 Then
            ldecDscto = ClsPanorama.FobjValorCampo(ldtbRes(0)(0), EnuTipoValor.EnuDecimal)
        End If
        Return ldecDscto
    End Function
    Private Shared Function FdecRevDsctosPredio(astrIdPredio As String) As Decimal
        Dim ldecDscto = 0D
        Dim lstrFechaIni As String = "'" & ClsPanoramaDat.FstrFechaNormalizada(
                GobjParametros.ObjAnoActual.DtmFechaInicioAno) & "'"
        Dim lstrFechaFin As String = "'" & ClsPanoramaDat.FstrFechaNormalizada(
                Date.Today) & "'"
        Dim lstrTablaPri = ClsNovedad.SstrNombreTabla
        Dim lstrTablaSec = ClsItemFactura.SstrNombreTabla
        Dim lstrCampSelPro As String() = {"SUM(" & ClsValor_NovDec.SstrNombreCampoBd & ")"}
        Dim lstrCampSelSec As String() = {}
        Dim lstrCampRelPri As String() = {StrCampoCarpeta, StrCampoCentroUtil,
                    ClsPrefijoFact_NovStr.SstrNombreCampoBd, ClsIdFactura_NovEnt.SstrNombreCampoBd,
                    ClsIdItemDocOrigen_NovShr.SstrNombreCampoBd}
        Dim lstrCampRelSec As String() = {StrCampoCarpeta, StrCampoCentroUtil,
                    ClsPrefijo_ItemFactStr.SstrNombreCampoBd,
                    ClsIdFactura_ItemFactEnt.SstrNombreCampoBd,
                    ClsIdItemFacturaShr.SstrNombreCampoBd}
        Dim lstrFiltro = StrFiltroUbicacion_Pri & " AND " & ClsIdPredio_ItemFactStr.SstrNombreCampoBd &
                " = '" & astrIdPredio & "' AND P." & ClsIdAno_NovShr.SstrNombreCampoBd & " = " &
                 Date.Today.Year & " AND " & ClsFechaNovedadDtm.SstrNombreCampoBd & " BETWEEN " &
                 lstrFechaIni & " AND " & lstrFechaFin & " AND " &
                 ClsIdTipoNovedadByt.SstrNombreCampoBd & " = " & EnuTipoNov.EnuRCrDctoCap
        Dim ldtbRes = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCampSelPro, lstrTablaSec,
                lstrCampSelSec, lstrCampRelPri, lstrCampRelSec, {{"", ""}}, True, lstrFiltro, {})
        If ldtbRes.Rows.Count > 0 Then
            ldecDscto = ClsPanorama.FobjValorCampo(ldtbRes(0)(0), EnuTipoValor.EnuDecimal)
        End If
        Return ldecDscto
    End Function
    ''' <summary>
    ''' Devuelve el valor descontado a los servicios de cuota de administración desde el primero del año
    ''' actual para tenerlo en cuenta en el calculo del ajuste cuando la nueva cuota es menor a lo que
    ''' se estaba conrando desde enros del año.Se utiliza solo cuando hay descuentos por pronto pago!
    ''' </summary>
    ''' <param name="astrIdPredio">Id del Predio para el cual se calcula el valor del descuento</param>
    ''' <returns></returns>
    Private Shared Function FdecSaldoDsctoPredio_NoPP(astrIdPredio As String) As Decimal
        Dim ldecDescuento = FdecDsctoPredio_NoPP(astrIdPredio)
        Dim ldecRevDescuento = FdecRevDsctoPredio_NoPP(astrIdPredio)
        Dim ldecSaldoDsctoNoPP = ldecDescuento - ldecRevDescuento
        If ldecSaldoDsctoNoPP < 0 Then
            Throw New ErrorInesperadoPanLException("Saldo descuento negativp!")
        End If
        Return ldecSaldoDsctoNoPP
    End Function
    Private Shared Function FdecDsctoPredio_NoPP(astrIdPredio As String) As Decimal
        Dim ldecDscto = 0D
        Dim ldtbDsctosPredio = FdtbDsctosPredio_Todos(astrIdPredio)
        Dim lenuTipoDocOri As EnuTipoDocOri, lstrPrefDocOri As String, lentIdDocOri As Integer,
                lshrIdItemDocOri As Short
        Dim lobjNotaCr As New ClsNotaCr(), lobjValorLlave As Object()
        Dim lobjItemNotaCr As ClsItemNotaCr
        For Each ldrwDscto As DataRow In ldtbDsctosPredio.Rows
            lenuTipoDocOri = ClsPanorama.FobjValorCampo(ldrwDscto(
                    ClsIdTipoDocOrigenByt.SstrNombreCampoBd), EnuTipoValor.EnuByte)
            lstrPrefDocOri = ClsPanorama.FobjValorCampo(ldrwDscto(
                    ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd), EnuTipoValor.EnuString)
            lentIdDocOri = ClsPanorama.FobjValorCampo(ldrwDscto(
                    ClsIdDocOrigenEnt.SstrNombreCampoBd), EnuTipoValor.EnuInteger)
            lshrIdItemDocOri = ClsPanorama.FobjValorCampo(ldrwDscto(
                    ClsIdItemDocOrigen_NovShr.SstrNombreCampoBd), EnuTipoValor.EnuShort)
            If lenuTipoDocOri = EnuTipoDocOri.EnuNotaCr Then
                lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrPrefDocOri,
                    lentIdDocOri}
                lobjNotaCr.SAbra(lobjValorLlave)
                lobjItemNotaCr = lobjNotaCr.ColItemsNotaCr(lshrIdItemDocOri.ToString())
                If lobjItemNotaCr.ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro =
                        EnuTipoDescuentoDef.EnuDsctoCapital Then
                    ldecDscto = ClsPanorama.FobjValorCampo(ldrwDscto(
                            ClsValor_NovDec.SstrNombreCampoBd), EnuTipoValor.EnuDecimal)
                End If
            End If
        Next
        Return ldecDscto
    End Function
    Private Shared Function FdecRevDsctoPredio_NoPP(astrIdPredio As String) As Decimal
        Dim ldecDsctoRev = 0D
        Dim ldtbRevDsctosPredio = FdtbRevDsctosPredio_Todos(astrIdPredio)
        Dim lenuTipoDocOri As EnuTipoDocOri, lstrPrefDocOri As String, lentIdDocOri As Integer,
                lshrIdItemDocOri As Short
        Dim lobjNotaRevCr As New ClsNotaReversionCr(), lobjValorLlave As Object()
        Dim lobjNotaCr As ClsNotaCr
        For Each ldrwDscto As DataRow In ldtbRevDsctosPredio.Rows
            lenuTipoDocOri = ClsPanorama.FobjValorCampo(ldrwDscto(
                    ClsIdTipoDocOrigenByt.SstrNombreCampoBd), EnuTipoValor.EnuByte)
            lstrPrefDocOri = ClsPanorama.FobjValorCampo(ldrwDscto(
                    ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd), EnuTipoValor.EnuString)
            lentIdDocOri = ClsPanorama.FobjValorCampo(ldrwDscto(
                    ClsIdDocOrigenEnt.SstrNombreCampoBd), EnuTipoValor.EnuInteger)
            lshrIdItemDocOri = ClsPanorama.FobjValorCampo(ldrwDscto(
                    ClsIdItemDocOrigen_NovShr.SstrNombreCampoBd), EnuTipoValor.EnuShort)
            If lenuTipoDocOri = EnuTipoDocOri.EnuNotaRevCr Then
                lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrPrefDocOri,
                        lentIdDocOri}
                lobjNotaRevCr.SAbra(lobjValorLlave)
                If lobjNotaRevCr.ObjTipoDocReversadoByt.ObjValorPro = EnuDocReversado.EnuNotaCr Then
                    lobjNotaCr = lobjNotaRevCr.ObjDocReversado
                    ldecDsctoRev += FdecDsctoRev(lobjNotaCr)
                End If
            End If
        Next
        Return ldecDsctoRev
    End Function
    Private Shared Function FdecDsctoRev(aobjNotaCr As ClsNotaCr) As Decimal
        Dim ldecDsctoRev = 0D
        For Each lobjItemNCr As ClsItemNotaCr In aobjNotaCr.ColItemsNotaCr
            If lobjItemNCr.ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro =
                    EnuTipoDescuentoDef.EnuDsctoPP Then
                ldecDsctoRev = lobjItemNCr.ObjValor_ItemNotaCrDec.ObjValorPro
            End If
        Next
        Return ldecDsctoRev
    End Function
    Private Shared Function FdtbDsctosPredio_Todos(astrIdPredio As String) As DataTable
        Dim lstrFechaIni As String = "'" & ClsPanoramaDat.FstrFechaNormalizada(
                GobjParametros.ObjAnoActual.DtmFechaInicioAno) & "'"
        Dim lstrFechaFin As String = "'" & ClsPanoramaDat.FstrFechaNormalizada(
                Date.Today) & "'"
        Dim lstrTablaPri = ClsNovedad.SstrNombreTabla
        Dim lstrTablaSec = ClsItemFactura.SstrNombreTabla
        Dim lstrCampSelPri As String() = {ClsIdTipoDocOrigenByt.SstrNombreCampoBd,
                ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd, ClsIdDocOrigenEnt.SstrNombreCampoBd,
                ClsIdItemDocOrigen_NovShr.SstrNombreCampoBd, ClsValor_NovDec.SstrNombreCampoBd}
        Dim lstrCampSelSec As String() = {}
        Dim lstrCampRelPri As String() = {StrCampoCarpeta, StrCampoCentroUtil,
                    ClsPrefijoFact_NovStr.SstrNombreCampoBd, ClsIdFactura_NovEnt.SstrNombreCampoBd,
                    ClsIdItemDocOrigen_NovShr.SstrNombreCampoBd}
        Dim lstrCampRelSec As String() = {StrCampoCarpeta, StrCampoCentroUtil,
                    ClsPrefijo_ItemFactStr.SstrNombreCampoBd,
                    ClsIdFactura_ItemFactEnt.SstrNombreCampoBd,
                    ClsIdItemFacturaShr.SstrNombreCampoBd}
        Dim lstrFiltro = StrFiltroUbicacion_Pri & " AND " & ClsIdPredio_ItemFactStr.SstrNombreCampoBd &
                " = '" & astrIdPredio & "' AND P." & ClsIdAno_NovShr.SstrNombreCampoBd & " = " &
                 Date.Today.Year & " AND " & ClsFechaNovedadDtm.SstrNombreCampoBd & " BETWEEN " &
                 lstrFechaIni & " AND " & lstrFechaFin & " AND " &
                 ClsIdTipoNovedadByt.SstrNombreCampoBd & " = " & EnuTipoNov.EnuCrDctoCap
        Dim ldtbRes = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCampSelPri, lstrTablaSec,
                lstrCampSelSec, lstrCampRelPri, lstrCampRelSec, {{"", ""}}, True, lstrFiltro, {})
        Return ldtbRes
    End Function
    Private Shared Function FdtbRevDsctosPredio_Todos(astrIdPredio As String) As DataTable
        Dim lstrFechaIni As String = "'" & ClsPanoramaDat.FstrFechaNormalizada(
                GobjParametros.ObjAnoActual.DtmFechaInicioAno) & "'"
        Dim lstrFechaFin As String = "'" & ClsPanoramaDat.FstrFechaNormalizada(
                Date.Today) & "'"
        Dim lstrTablaPri = ClsNovedad.SstrNombreTabla
        Dim lstrTablaSec = ClsItemFactura.SstrNombreTabla
        Dim lstrCampSelPri As String() = {ClsIdTipoDocOrigenByt.SstrNombreCampoBd,
                ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd, ClsIdDocOrigenEnt.SstrNombreCampoBd,
                ClsIdItemDocOrigen_NovShr.SstrNombreCampoBd, ClsValor_NovDec.SstrNombreCampoBd}
        Dim lstrCampSelSec As String() = {}
        Dim lstrCampRelPri As String() = {StrCampoCarpeta, StrCampoCentroUtil,
                    ClsPrefijoFact_NovStr.SstrNombreCampoBd, ClsIdFactura_NovEnt.SstrNombreCampoBd,
                    ClsIdItemDocOrigen_NovShr.SstrNombreCampoBd}
        Dim lstrCampRelSec As String() = {StrCampoCarpeta, StrCampoCentroUtil,
                    ClsPrefijo_ItemFactStr.SstrNombreCampoBd,
                    ClsIdFactura_ItemFactEnt.SstrNombreCampoBd,
                    ClsIdItemFacturaShr.SstrNombreCampoBd}
        Dim lstrFiltro = StrFiltroUbicacion_Pri & " AND " & ClsIdPredio_ItemFactStr.SstrNombreCampoBd &
                " = '" & astrIdPredio & "' AND P." & ClsIdAno_NovShr.SstrNombreCampoBd & " = " &
                 Date.Today.Year & " AND " & ClsFechaNovedadDtm.SstrNombreCampoBd & " BETWEEN " &
                 lstrFechaIni & " AND " & lstrFechaFin & " AND " &
                 ClsIdTipoNovedadByt.SstrNombreCampoBd & " = " & EnuTipoNov.EnuRCrDctoCap
        Dim ldtbRes = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCampSelPri, lstrTablaSec,
                lstrCampSelSec, lstrCampRelPri, lstrCampRelSec, {{"", ""}}, True, lstrFiltro, {})
        Return ldtbRes
    End Function
    Private Shared Function FdtbValorNotasAjuste() As DataTable
        Dim lshrIdAno As Short = GobjParametros.ObjAnoActual.ObjIdAnoShr.ObjValorPro
        Dim lstrTabla = ClsNotaAjusteCuotaAdmin.SstrNombreTabla
        Dim lstrFechaIni As String = " '" & ClsPanoramaDat.FstrFechaNormalizada(DateSerial(
                Date.Today.Year, 1, 1)) & "'"
        Dim lstrFecha As String = "'" & ClsPanoramaDat.FstrFechaNormalizada(Date.Today) & "'"
        Dim lstrCamSel As String() = {ClsIdPredio_NotaAjusteStr.SstrNombreCampoBd,
                "SUM(" & ClsValor_NotaAjusteDec.SstrNombreCampoBd & ")"}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " &
                ClsFecha_NotaAjusteDtm.SstrNombreCampoBd & " Between " & lstrFechaIni & " AND " &
                lstrFecha
        Dim lstrGrupo As String() = {ClsIdPredio_NotaAjusteStr.SstrNombreCampoBd}
        Dim lstrIndice As String(,) = {{ClsIdPredio_NotaAjusteStr.SstrNombreCampoBd, "ASC"}}
        Dim ldtbVlrAjus = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamSel, lstrIndice,
                lstrFiltro, True, lstrGrupo)
        Return ldtbVlrAjus
    End Function
    Private Shared Function FdecValorDebioFacturarAPredio(aobjServicio As ClsServicio,
            astrPeriodoFin As String, astrIdPredio As String) As Decimal
        Dim ldtbValorDefiPeriodoPredio As DataTable
        Dim lshrIdAno As Short = aobjServicio.ObjMiAno.ObjIdAnoShr.ObjValorPro
        Dim lstrPeriodoIni As String
        Dim lstrPeriodoIniGen = String.Empty
        Dim ldecValor = 0D
        If GobjParametros.ColAnos.Count = 1 Then
            lstrPeriodoIniGen = lshrIdAno.ToString() & "01"
        End If
        If Not String.IsNullOrEmpty(lstrPeriodoIniGen) Then
            lstrPeriodoIni = lstrPeriodoIniGen
        Else
            lstrPeriodoIni = FstrPeriodoIniFact(lshrIdAno, astrIdPredio)
        End If
        If Not String.IsNullOrEmpty(lstrPeriodoIni) Then
            Dim lentCantPeriodos As Integer = CType(Right(astrPeriodoFin, 2), Integer) -
                    CType(Right(lstrPeriodoIni, 2), Integer)
            ldtbValorDefiPeriodoPredio = FdtbValorDefiPeriodoPredio(aobjServicio, astrIdPredio)
            If ldtbValorDefiPeriodoPredio.Rows.Count > 0 Then
                Dim ldrwPredio = ldtbValorDefiPeriodoPredio.Rows(0)
                ldecValor = ClsPanorama.FobjValorCampo(ldrwPredio(
                        ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd),
                        EnuTipoValor.EnuDecimal)
                ldecValor *= lentCantPeriodos
            End If
        End If
        Return ldecValor
    End Function
    ''' <summary>
    ''' Devuelve periodo de inicio de facturación en el año y al predio pasados en los argumentos
    ''' </summary>
    ''' <returns></returns>
    Private Shared Function FstrPeriodoIniFact(ashrIdAno As Short, astrIdpredio As String) As String
        Dim lstrTabla = ClsItemFactura.SstrNombreTabla
        Dim lstrCamSel = {"MIN(" & ClsPeriodo_ItemFactStr.SstrNombreCampoBd & ") AS Periodo"}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " &
                ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd & " = " & ashrIdAno &
                " AND " & ClsIdPredio_ItemFactStr.SstrNombreCampoBd &
                " = '" & astrIdpredio & "'"
        Dim ldtbPer = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamSel, {{"", ""}}, lstrFiltro)
        Dim lstrPeriodo = ClsPanorama.FobjValorCampo(ldtbPer(0)("Periodo"), EnuTipoValor.EnuString)
        Return lstrPeriodo
    End Function
    Private Shared Function FstrPeriodoIniFact() As String
        Dim lstrTablaPri = ClsFactura.SstrNombreTabla
        Dim lstrTablaSec = ClsItemFactura.SstrNombreTabla
        Dim lstrCamSelPri As String() = {}
        Dim lstrCamSelSec As String() = {"MIN(" & ClsPeriodo_ItemFactStr.SstrNombreCampoBd & ")"}
        Dim lstrCamRelPri As String() = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsPrefijo_FactStr.SstrNombreCampoBd, ClsIdFacturaEnt.SstrNombreCampoBd}
        Dim lstrCamRelSec As String() = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsPrefijo_ItemFactStr.SstrNombreCampoBd, ClsIdFactura_ItemFactEnt.SstrNombreCampoBd}
        Dim lstrFiltro As String = StrFiltroUbicacion_Pri & " AND " &
                ClsIdModoFacturacionByt.SstrNombreCampoBd & " = " & EnuModoFacturacionDef.EnuSistema
        Dim lstrGrupo As String() = {}
        Dim lstrOrden As String(,) = {{"", ""}}
        Dim ldtbRes = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCamSelPri, lstrTablaSec,
                lstrCamSelSec, lstrCamRelPri, lstrCamRelSec, lstrOrden, True, lstrFiltro, lstrGrupo)
        Dim lstrPeriodo = ClsPanorama.FobjValorCampo(ldtbRes(0)(0), EnuTipoValor.EnuString)
        Return lstrPeriodo
    End Function
    ''' <summary>
    ''' Devuelve la cantidad de periodos facturados antes del periodo actuaL
    ''' </summary>
    ''' <returns></returns>
    Private Function FentCanPeriodosFacturados(aobjAno As ClsAno) As Integer
        Dim lstrPeriodoIni = FstrPeriodoIniFact()
        Dim lstrPeriodoFin = aobjAno.StrIdPeriodoActual
        Dim lentCantPeriodos As Integer = CType(Right(lstrPeriodoFin, 2), Integer) -
                    CType(Right(lstrPeriodoIni, 2), Integer)
        If lentCantPeriodos > 12 Then
            Throw New PanLException("No corresponde al primer año de la aplicación!")
        End If
        Return lentCantPeriodos
    End Function
    Private Shared Function FdtbValorDefiPeriodoPredio(aobjServicio As ClsServicio,
            astrIdPredio As String) As DataTable
        Dim lshrIdServicio As Short = aobjServicio.ObjIdServicioShr.ObjValorPro
        Dim lshrIdAno As Short = aobjServicio.ObjIdAno_ServicioShr.ObjValorPro
        Dim lstrCamposSelect = {ClsIdPredio_ItemProgramaFactStr.SstrNombreCampoBd,
                ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd}
        Dim lstbExpSql As New StringBuilder
        ' Filtro
        With lstbExpSql
            .Append(StrFiltroUbicacion).Append(" AND ").Append(ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd)
            .Append(" = ").Append(lshrIdAno).Append(" AND ")
            .Append(ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd).Append(" = ")
            .Append(lshrIdServicio).Append(" AND ")
            .Append(ClsIdPredio_ItemProgramaFactStr.SstrNombreCampoBd).Append(" = '")
            .Append(astrIdPredio).Append("'")
        End With
        Dim lstrFiltro = lstbExpSql.ToString
        Dim lstrIndice(,) = {{ClsIdPredio_ItemFactStr.SstrNombreCampoBd, "ASC"}}
        Dim ldtbValorDefiPeriodoPredios = ClsPanorama.FdtbDataTable(ClsItemProgramaFact.SstrNombreTabla,
lstrCamposSelect, lstrIndice, lstrFiltro)
        Return ldtbValorDefiPeriodoPredios
    End Function
    Private Shared Sub SCreeItemProgramaFactPredio(aobjServicio As ClsServicio,
            adecValorPeriodo As Decimal, aobjPredio As ClsPredio)
        Dim ldrwNewItemProgramaFact = FdrwNewItemProgramaFact()
        Dim lobjItemProgramFact As New ClsItemProgramaFact(aobjPredio,
                EnuTipoDeudorDef.EnuPredio, ldrwNewItemProgramaFact)
        With lobjItemProgramFact
            .SCreeObj(Nothing)
            .ObjIdCarpeta_ItemProgramaFactShr.ObjValorPro = GshrIdCarpeta
            .ObjIdCentroUtil_ItemProgramaFactShr.ObjValorPro = GshrIdCentroUtil
            .ObjIdAno_ItemProgramaFactShr.ObjValorPro = aobjServicio.ObjIdAno_ServicioShr.ObjValorPro
            .ObjIdServicio_ItemProgramaFactShr.ObjValorPro = aobjServicio.ObjIdServicioShr.ObjValorPro
            .ObjIdCliente_ItemProgramaFactDbl.ObjValorPro = 0
            .ObjIdPredio_ItemProgramaFactStr.ObjValorPro = aobjPredio.ObjIdPredioStr.ObjValorPro
            .ObjCantidadPeriodosShr.ObjValorPro = aobjServicio.ObjCantPeriodos_ServicioShr.ObjValorPro
            .ObjPeriodoIni_ItemProgStr.ObjValorPro = aobjServicio.ObjPeriodoInicioStr.ObjValorPro
            .ObjOrigen_ItemProgramaFacByt.ObjValorPro = EnuOrigenItemProgramaFactDef.EnuAplicacion
            .ObjValorPeriodo_ItemProgramaFactDec.ObjValorPro = adecValorPeriodo
            .SActualice(True)
        End With
    End Sub
    Private Shared Function FdrwNewItemProgramaFact() As DataRow
        Dim ldtbItemsProgramaFac As DataTable = Nothing
        If IsNothing(ldtbItemsProgramaFac) Then
            Dim lstrFiltro = StrFiltroUbicacion & " AND " & ClsOrdinal_ItemProgramaFact.SstrNombreCampoBd & " = 0"
            ldtbItemsProgramaFac = ClsPanorama.FdtbDataTable(ClsItemProgramaFact.SstrNombreTabla, {"*"},
                            {{"", ""}}, lstrFiltro)
        End If
        Dim ldrwItemProg = ldtbItemsProgramaFac.NewRow
        Return ldrwItemProg
    End Function
    Friend Shared Function FdecValorAjuste(aobjServicioAjuste As ClsServicio) As Decimal
        If Not aobjServicioAjuste.ObjEsAjusteBln.ObjValorPro Then
            Throw New ErrorInesperadoPanLException("Es Servicio no es de Ajuste")
        End If
        Dim lshrIdServicio As Short = aobjServicioAjuste.ObjIdServicioShr.ObjValorPro
        Dim lshrIdAno As Short = aobjServicioAjuste.ObjIdAno_ServicioShr.ObjValorPro
        ' Valor Cuotas de Ajuste
        Dim ldecValorCuotasAjuste = FdecValorCuotasAjuste(lshrIdAno, lshrIdServicio)
        ldecValorCuotasAjuste *= aobjServicioAjuste.ObjCantPeriodos_ServicioShr.ObjValorPro
        ' Valor Anticipos Generados
        Dim ldecValorAnticiposAjuste = FdecValorAnticiposAjuste(lshrIdAno)
        Dim ldecValorAjuste = ldecValorCuotasAjuste - ldecValorAnticiposAjuste
        Return ldecValorAjuste
    End Function
    Private Shared Function FdecValorCuotasAjuste(ashrIdAnoAjuste As Short, ashrIdServicioAjuste As Short)
        Dim lstbExpSql As New StringBuilder
        ' CampoSelect
        With lstbExpSql
            .Clear.Append("SUM(").Append(ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd).Append(")")
        End With
        Dim lstrCamposSelect = {lstbExpSql.ToString}
        ' Filtro
        With lstbExpSql
            .Clear.Append(StrFiltroUbicacion).Append(" AND ").Append(ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd)
            .Append(" = ").Append(ashrIdAnoAjuste).Append(" AND ")
            .Append(ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd).Append(" = ")
            .Append(ashrIdServicioAjuste)
        End With
        Dim lstrFiltro = lstbExpSql.ToString
        Dim lstrIndice(,) = {{"", ""}}
        Dim ldtbValorPeriodo = ClsPanorama.FdtbDataTable(ClsItemProgramaFact.SstrNombreTabla,
lstrCamposSelect, lstrIndice, lstrFiltro)
        Dim ldecValorPeriodo As Decimal = ClsPanorama.FobjValorCampo(ldtbValorPeriodo.Rows(0)(0),
EnuTipoValor.EnuDecimal)
        Return ldecValorPeriodo
    End Function
    Private Shared Function FdecValorAnticiposAjuste(ashrIdAno As Short) As Decimal
        Dim lstbExpSql As New StringBuilder
        ' CampoSelect
        With lstbExpSql
            .Clear.Append("SUM(").Append(ClsValor_AntDec.SstrNombreCampoBd).Append(")")
        End With
        Dim lstrCamposSelect = {lstbExpSql.ToString}
        ' Filtro
        With lstbExpSql
            .Clear.Append(StrFiltroUbicacion).Append(" AND ").Append(ClsIdTipoDocOrigen_AntByt.SstrNombreCampoBd)
            .Append(" = ").Append(EnuTipoDocOri.EnuNotaAjuste).Append(" AND ")
            .Append("YEAR(").Append(ClsFechaAnticipoDtm.SstrNombreCampoBd).Append(") = ")
            .Append(ashrIdAno)
        End With
        Dim lstrFiltro = lstbExpSql.ToString
        Dim lstrIndice(,) = {{"", ""}}
        Dim ldtbValorAnticiposAjuste = ClsPanorama.FdtbDataTable(ClsAnticipo.SstrNombreTabla,
                lstrCamposSelect, lstrIndice, lstrFiltro)
        Dim ldecValorAnticiposAjuste As Decimal = ClsPanorama.FobjValorCampo(ldtbValorAnticiposAjuste.Rows(0)(0),
                EnuTipoValor.EnuDecimal)
        Return ldecValorAnticiposAjuste
    End Function
    'Ajuste descuento pronto pago
    Friend Shared Function FdecAjusteDsctoPP(aobjPredioAgr As ClsPredio,
            adecValorItemAjuste As Decimal) As Decimal
        Dim ldecAjusteDsctoPP = 0D, lentCantPerioRR = 0, lentCantPerioNC = 0
        Dim lstrIdPred = aobjPredioAgr.ObjIdPredioStr.ObjValorPro
        Dim ldecVlrDsctoPPAplicado = FdecValorDsctoPPEnRecCaja(lstrIdPred, lentCantPerioRR) +
                FdecValorDsctoPPEnAnt(lstrIdPred, lentCantPerioNC)
        Dim lentCantPerioConDectoPP = lentCantPerioRR + lentCantPerioNC
        If ldecVlrDsctoPPAplicado > 0 Then
            If GobjParametros.ObjAnoActual.ObjTipoDsctoPPByt.ObjValorPro =
                    EnuTipoDsctoPP.EnuValorFijo Then
                Dim ldecValorDsctoActual = FdecVlrFijoDsctoActual(aobjPredioAgr)
                ldecAjusteDsctoPP = (ldecValorDsctoActual * lentCantPerioConDectoPP) - ldecVlrDsctoPPAplicado
            Else
                ldecAjusteDsctoPP = FdecVlrAjusteDsctoPPPorciento(aobjPredioAgr, adecValorItemAjuste)
            End If
        End If
        Return ldecAjusteDsctoPP
    End Function
    Private Shared Function FdecValorDsctoPPEnRecCaja(astrIdpredioAgr As String,
            ByRef aentCantPeriodos As Integer) As Decimal
        Dim lstrFecha = "'" & ClsPanoramaDat.FstrFechaNormalizada(
                GobjParametros.ObjAnoActual.DtmFechaInicioAno) & "'"
        Dim lstrTablaPri = ClsItemRecCaja.SstrNombreTabla
        Dim lstrCamposPri = {ClsValor_ItemRecDec.SstrNombreCampoBd}
        Dim lstrTablaSec = ClsReciboCaja.SstrNombreTabla
        Dim lstrCamposSec = {ClsIdPredioAgrupador_RecStr.SstrNombreCampoBd}
        Dim lstrCamposRelPri = {StrCampoCarpeta, StrCampoCentroUtil,
ClsPrefijo_RecStr.SstrNombreCampoBd,
ClsIdRecCaja_ItemRecEnt.SstrNombreCampoBd}
        Dim lstrCamposRelSec = {StrCampoCarpeta, StrCampoCentroUtil,
ClsPrefijo_RecStr.SstrNombreCampoBd,
ClsIdRecCajaEnt.SstrNombreCampoBd}
        Dim lstrFiltro = StrFiltroUbicacion_Pri & " AND " &
                    ClsIdPredioAgrupador_RecStr.SstrNombreCampoBd & " = '" &
                    astrIdpredioAgr & "' AND " & ClsIdTipoItemRecByt.SstrNombreCampoBd &
                    " = " & EnuTipoItemRecCajaDef.EnuDsctoPP & " AND " &
                    ClsFechaRecDtm.SstrNombreCampoBd & " >= " & lstrFecha & " AND P." &
                    ClsValor_ItemRecDec.SstrNombreCampoBd & " > 0"
        Dim ldtbValoresDsctoPP = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCamposPri, lstrTablaSec,
                lstrCamposSec, lstrCamposRelPri, lstrCamposRelSec, {{"", ""}}, lstrFiltro,
                Array.Empty(Of String)(), True)
        Dim ldecVlrDsctoPPAplicado As Decimal
        Dim ldecVlrTotalDsctoPP = 0D
        For Each ldrwDsctoPP As DataRow In ldtbValoresDsctoPP.Rows
            ldecVlrDsctoPPAplicado = ClsPanorama.FobjValorCampo(ldrwDsctoPP(
                    ClsValor_ItemRecDec.SstrNombreCampoBd), EnuTipoValor.EnuDecimal)
            If ldecVlrDsctoPPAplicado > 0 Then
                ldecVlrTotalDsctoPP += ldecVlrDsctoPPAplicado
                aentCantPeriodos += 1
            End If
        Next
        Return ldecVlrTotalDsctoPP
    End Function
    Private Shared Function FdecValorDsctoPPEnAnt(astrIdpredioAgr As String,
            ByRef aentCantPeriodos As Integer)
        Dim lstrFecha = "'" & ClsPanoramaDat.FstrFechaNormalizada(GobjParametros.ObjAnoActual.DtmFechaInicioAno) & "'"
        Dim lstrTablaPri = ClsItemNotaCon.SstrNombreTabla
        Dim lstrCamposPri = {ClsValor_NotaConDec.SstrNombreCampoBd}
        Dim lstrTablaSec = ClsNotaCon.SstrNombreTabla
        Dim lstrCamposSec = {ClsIdPredioAgrupador_NotaConStr.SstrNombreCampoBd}
        Dim lstrCamposRelPri = {StrCampoCarpeta, StrCampoCentroUtil,
ClsPrefijo_NotaConStr.SstrNombreCampoBd,
ClsIdNotaCon_ItemNotaConEnt.SstrNombreCampoBd}
        Dim lstrCamposRelSec = {StrCampoCarpeta, StrCampoCentroUtil,
ClsPrefijo_NotaConStr.SstrNombreCampoBd,
ClsIdNotaConEnt.SstrNombreCampoBd}
        Dim lstrFiltro = StrFiltroUbicacion_Pri & " AND " &
                ClsIdPredioAgrupador_NotaConStr.SstrNombreCampoBd & " = '" &
                astrIdpredioAgr & "' AND " & ClsIdTipoItemNotaConByt.SstrNombreCampoBd &
                " = " & EnuTipoItemNotaConDef.EnuDsctoPP & " AND " &
                ClsFecha_NotaConDtm.SstrNombreCampoBd & " >= " & lstrFecha & " AND P." &
                ClsValor_ItemNotaConDec.SstrNombreCampoBd & " > 0"
        Dim ldtbValoresDsctoPP = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCamposPri, lstrTablaSec,
                lstrCamposSec, lstrCamposRelPri, lstrCamposRelSec, {{"", ""}}, lstrFiltro,
                Array.Empty(Of String)(), True)
        Dim ldecVlrDsctoPPAplicado As Decimal
        Dim ldecVlrTotalDsct = 0D
        For Each ldrwDsctoPP As DataRow In ldtbValoresDsctoPP.Rows
            ldecVlrDsctoPPAplicado = ClsPanorama.FobjValorCampo(ldrwDsctoPP(
ClsValor_NotaConDec.SstrNombreCampoBd), EnuTipoValor.EnuDecimal)
            If ldecVlrDsctoPPAplicado > 0 Then
                ldecVlrTotalDsct += ldecVlrDsctoPPAplicado
                aentCantPeriodos += 1
            End If
        Next
        Return ldecVlrTotalDsct
    End Function
    Private Shared Function FdecVlrFijoDsctoActual(aobjPredio As ClsPredio) As Decimal
        Dim lobjSectorPredio As ClsSector = GobjParametros.ColSectores(
                aobjPredio.ObjIdSector_PredioShr.ToString)
        Dim ldblMontoDsctoPP As Double = lobjSectorPredio.ObjDctoProntoPago_SecDbl.ObjValorPro
        Return CType(ldblMontoDsctoPP, Decimal)
    End Function
    Private Shared Function FdecVlrAjusteDsctoPPPorciento(aobjPredio As ClsPredio,
            adecValorItemFra As Decimal)
        Dim lobjSectorPredio As ClsSector = GobjParametros.ColSectores(
                aobjPredio.ObjIdSector_PredioShr.ToString)
        Dim ldblMontoDsctoPP As Double = lobjSectorPredio.ObjDctoProntoPago_SecDbl.ObjValorPro
        Dim ldecAjusteDscto As Decimal = adecValorItemFra * ldblMontoDsctoPP
        ldecAjusteDscto = FdecValorRedondeado(ldecAjusteDscto)
        Return ldecAjusteDscto
    End Function
    ' Supresión servicio de ajuste generado en el periodo actual
    Private Shared Sub SSuprimaSerAjuste(aobjSerAjuste As ClsServicio)
        ' Suprimo los items del programa de facturación
        SSuprimaItemsProgFac(aobjSerAjuste)
        ' Suprimo Notas de Ajuste Generadas
        SSuprimoNotasAjuste(aobjSerAjuste)
        ' Suprim Servicio
        If Not aobjSerAjuste.EnuPermisosObj And EnuPermisosDef.EnuSuprimir Then
            aobjSerAjuste.EnuPermisosObj += EnuPermisosDef.EnuSuprimir
        End If
        Dim lstrKey = aobjSerAjuste.ObjIdAno_ServicioShr.ToString & "," &
aobjSerAjuste.ObjIdServicioShr.ToString
        aobjSerAjuste.ObjMiAno.SRemuevaServicio(lstrKey)
        aobjSerAjuste.FblnSuprimio()
    End Sub
    Private Shared Sub SSuprimaItemsProgFac(aobjServicio As ClsServicio)
        Dim lstrTabla = ClsItemProgramaFact.SstrNombreTabla
        Dim lcolCamposRef As New Collection From {
            StrCampoCarpeta,
            StrCampoCentroUtil,
            ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd,
            ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd
        }
        Dim lcolDatosRef As New Collection From {
                {GshrIdCarpeta, StrCampoCarpeta},
                {GshrIdCentroUtil, StrCampoCentroUtil},
                {aobjServicio.ObjIdAno_ServicioShr.ObjValorPro,
                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd},
                {aobjServicio.ObjIdServicioShr.ObjValorPro,
                ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd}
        }
        GobjPanDat.SElimineRegistro(lstrTabla, lcolCamposRef, lcolDatosRef)
    End Sub
    Private Shared Sub SSuprimoNotasAjuste(aobjServicio As ClsServicio)
        ' Se suprimen las notas de ajuste generadas en el período actual
        ' Suprimo las novedades
        Dim lstrTabla = ClsNovedadAnticipo.SstrNombreTabla
        Dim ldtmFecIni = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo()
        Dim lstrFecIni = ClsPanoramaDat.FstrFechaNormalizada(ldtmFecIni)
        Dim lstbSql As New StringBuilder
        With lstbSql
            .Append("DELETE FROM ").Append(lstrTabla).Append(" WHERE ")
            .Append(StrFiltroUbicacion).Append(" AND ")
            .Append(ClsIdTipoDocOrigen_NovAntByt.SstrNombreCampoBd).Append(" = ")
            .Append(EnuTipoDocOri.EnuNotaAjuste).Append(" AND ")
            .Append(ClsFechaNovedadAntDtm.SstrNombreCampoBd).Append(" >= '")
            .Append(lstrFecIni).Append("'")
        End With
        GobjPanDat.SEjecuteSentenciaSql(lstbSql.ToString())
        ' Suprimo NotasAjuste
        lstrTabla = ClsNotaAjusteCuotaAdmin.SstrNombreTabla
        With lstbSql
            .Clear.Append("DELETE FROM ").Append(lstrTabla).Append(" WHERE ")
            .Append(StrFiltroUbicacion).Append(" AND ")
            .Append(ClsFecha_NotaAjusteDtm.SstrNombreCampoBd).Append(" >= '")
            .Append(lstrFecIni).Append("'")
        End With
        GobjPanDat.SEjecuteSentenciaSql(lstbSql.ToString())
        ' suprimo los anticipos generados por las notas de ajuste
        lstrTabla = ClsAnticipo.SstrNombreTabla
        With lstbSql
            .Clear.Append("DELETE FROM ").Append(lstrTabla).Append(" WHERE ")
            .Append(StrFiltroUbicacion).Append(" AND ")
            .Append(ClsIdTipoDocOrigen_AntByt.SstrNombreCampoBd).Append(" = ")
            .Append(EnuTipoDocOri.EnuNotaAjuste).Append(" AND ")
            .Append(ClsFechaAnticipoDtm.SstrNombreCampoBd).Append(" >= '")
            .Append(lstrFecIni).Append("'")
        End With
        GobjPanDat.SEjecuteSentenciaSql(lstbSql.ToString())
    End Sub
#End Region
#Region "Procedimientos y funciones publicos"
#Region "Varios"
    Friend Shared Sub SInstancieCentroUtilOriCop()
        Dim ldtbCenutilOrionCop = ClsPanorama.FdtbDataTable(ClsCentroUtilOriCop.SstrNombreTabla,
                {"*"}, Nothing, StrFiltroUbicacion)
        If ldtbCenutilOrionCop.Rows.Count = 0 Then
            Dim ldrwCenUtiOriCop As DataRow = ldtbCenutilOrionCop.NewRow
            GobjParametros = New ClsCentroUtilOriCop(ldrwCenUtiOriCop)
            If Not CType(GobjParametros.EnuPermisosObj And EnuPermisosDef.EnuCrear, Boolean) Then
                GobjParametros.EnuPermisosObj += EnuPermisosDef.EnuCrear
            End If
            GobjParametros.SCreeObj(Nothing)
        Else
            GobjParametros = New ClsCentroUtilOriCop(ldtbCenutilOrionCop.Rows(0))
            GobjParametros.SInicialiceObj()
            GobjParametros.SLeaValores(True)
            GblnNotiSonoras = GobjParametros.ObjNotificacionesSonorasBln.ObjValorPro
        End If
        If GenuTipoInstanciamiento = EnuTipoInstanciamiento.enuActualizacion AndAlso
                GstrVerAntApp < "16.35.400.1306" Then
            SActualiceTotalAreaCopr()
        End If
    End Sub
    ''' <summary>
    ''' Devuelve la tasa efectiva anual para un valor en porcentaje seguida de dos caracteres que indican el primero
    ''' la periodicidad y el segundo el modo de pago (Vnecido o anticipado)
    ''' </summary>
    ''' <param name="astrValor">Un texto que contiene la tasa y el calificador</param>
    ''' <param name="aenuTipoInteres">In enumerador que indica si la tasa es simple o compuesta</param>
    ''' <returns>La tasa efectiva anual. Si alguno de los argumentos no es valido devuelve 0</returns>
    ''' <remarks>Los calificadores validos son: dv(día vencido), mv(mes vencido), bv(bimestre vencido, tv(trimestre vencido),
    ''' sv(semestre vencido), da(dia anticipado), ma(mes anticipado), ba(bimestre anticipado), ta(trimestre anticipado) y
    ''' sa(semestre anticipado)</remarks>
    Shared Function FdblTraduceATasaEfectivaAnual(astrValor As String,
            aenuTipoInteres As EnuTipoInteres) As Double
        If String.IsNullOrEmpty(astrValor) Then astrValor = String.Empty
        Dim lstrCalificadores = "dvmvbvtvsvdamabatasa"
        Dim ldblTasaEA As Double = 0
        Dim ldblValor As Double = Val(astrValor) / 100
        Dim lstrValor = String.Empty
        Dim lenuModoPagoIntereses = EnuModoPagoIntereses.None
        Dim i = 0
        For Each lchrCar As Char In astrValor
            If lchrCar <> "." AndAlso (lchrCar < "0" OrElse lchrCar > "9") Then
                i += 1
            End If
        Next
        If i <> 2 Then
            Return 0
        End If
        If lstrCalificadores.ToUpper.Contains(Right(astrValor.ToUpper, 2)) Then
            lstrValor = Right(astrValor.ToUpper, 2)
        End If
        If Not String.IsNullOrEmpty(lstrValor) Then
            If lstrValor.ToUpper.EndsWith("V") Then
                lenuModoPagoIntereses = EnuModoPagoIntereses.EnuVencido
            ElseIf UCase(Mid(lstrValor, 2, 1)) = "A" Then
                lenuModoPagoIntereses = EnuModoPagoIntereses.EnuAnticipado
            End If
            Select Case UCase(Mid(lstrValor, 1, 1))
                Case "D"
                    ldblTasaEA = FdblTasaEA(ldblValor, EnuPeriodicidadDePagoDef.EnuDiaria,
                            lenuModoPagoIntereses, aenuTipoInteres) * 100
                Case "M"
                    ldblTasaEA = FdblTasaEA(ldblValor, EnuPeriodicidadDePagoDef.EnuMensual,
                            lenuModoPagoIntereses, aenuTipoInteres) * 100
                Case "B"
                    ldblTasaEA = FdblTasaEA(ldblValor, EnuPeriodicidadDePagoDef.EnuBimestral,
                            lenuModoPagoIntereses, aenuTipoInteres) * 100
                Case "T"
                    ldblTasaEA = FdblTasaEA(ldblValor, EnuPeriodicidadDePagoDef.EnuTrimestral,
                            lenuModoPagoIntereses, aenuTipoInteres) * 100
                Case "S"
                    ldblTasaEA = FdblTasaEA(ldblValor, EnuPeriodicidadDePagoDef.EnuSemestral,
                            lenuModoPagoIntereses, aenuTipoInteres) * 100
            End Select
        End If
        Return ldblTasaEA
    End Function
    'devuelve la tasa efectiva anual a partir de la tasa de un periodo
    Shared Function FdblTasaEA(adblPeriodRate As Double, aenuPeriodicity As EnuPeriodicidadDePagoDef,
            aenuModoPago As EnuModoPagoIntereses, aenuInterestType As EnuTipoInteres) As Double
        Dim ldblCanPeriodos As Double = FsngPeriodosAño(aenuPeriodicity)
        Dim ldblTasaEA As Double
        If aenuInterestType = EnuTipoInteres.EnuInteresCompuesto Then
            If aenuModoPago = EnuModoPagoIntereses.EnuVencido Then
                ldblTasaEA = ((1 + adblPeriodRate) ^ ldblCanPeriodos) - 1
            Else
                ldblTasaEA = ((1 - adblPeriodRate) ^ (-ldblCanPeriodos)) - 1
            End If
        Else
            ldblTasaEA = adblPeriodRate * ldblCanPeriodos
        End If
        Return ldblTasaEA
    End Function
    ''' <summary>
    ''' Devuelve la cantidad de periodos que contiene un año de una periodicidad determinada por el
    ''' argumento #aenuPeriodicidad".
    ''' </summary>
    ''' <param name="aenuPeriodicidad">Indica la periodicidad a evaluar.</param>
    ''' <remarks></remarks>
    Shared Function FsngPeriodosAño(aenuPeriodicidad As EnuPeriodicidadDePagoDef) As Single
        Dim lsngPeriodos As Single
        Select Case aenuPeriodicidad
            Case EnuPeriodicidadDePagoDef.EnuDiaria
                lsngPeriodos = 360
            Case EnuPeriodicidadDePagoDef.EnuMensual
                lsngPeriodos = 12
            Case EnuPeriodicidadDePagoDef.EnuBimestral
                lsngPeriodos = 6
            Case EnuPeriodicidadDePagoDef.EnuTrimestral
                lsngPeriodos = 4
            Case EnuPeriodicidadDePagoDef.EnuSemestral
                lsngPeriodos = 2
            Case EnuPeriodicidadDePagoDef.EnuAnual
                lsngPeriodos = 1
            Case Else
                Throw New ValorArgumentoInvalidoException("El argumento 'aenuPeriodicidad' no esvalido")
        End Select
        Return lsngPeriodos
    End Function
    ''' <summary>
    ''' Devuelve un array de string indicando las diferencias entre las preferencias del predio 
    ''' agrupador y los predios agrupados
    ''' </summary>
    Friend Shared Function FstrPrediosDifieren() As String()
        Dim lcolPropPredAgr As Collection, lcolPropPred As Collection
        Dim lobjPropPredAgr As ClsPropietario
        Dim lstrDifsPredio() As String = Array.Empty(Of String)()
        Dim ldtbPrediosAgr = FdtbPrediosAgrupadores()
        Dim i = 0
        Dim lobjPredioAgru As New ClsPredio(EnuModoInstanciaObjDef.EnuUnico)
        Dim lstrIdPreAgr As String
        For Each ldrwPredioAgru As DataRow In ldtbPrediosAgr.Rows
            lstrIdPreAgr = ClsPanorama.FobjValorCampo(ldrwPredioAgru(0),
                    EnuTipoValor.EnuString)
            lobjPredioAgru.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrIdPreAgr})
            lcolPropPredAgr = lobjPredioAgru.ColPropietarios
            Dim lstrIdCliente As String
            Dim ldblPorParCli As Double
            For Each lobjPre As ClsPredio In lobjPredioAgru.ColPrediosAgrupados
                If lobjPre.ObjIdPredioStr.ObjValorPro <>
                        lobjPredioAgru.ObjIdPredioStr.ObjValorPro Then
                    lcolPropPred = lobjPre.ColPropietarios
                    If lcolPropPredAgr.Count = lcolPropPred.Count Then
                        For Each lobjProp As ClsPropietario In lcolPropPredAgr
                            lstrIdCliente = lobjProp.ObjIdCliente_PropDbl.ToString()
                            If lcolPropPredAgr.Contains(lstrIdCliente) Then
                                lobjPropPredAgr = lcolPropPredAgr(lstrIdCliente)
                                ldblPorParCli =
                                        lobjProp.ObjPorcentajePartiDbl.ObjValorPro
                                If lobjPropPredAgr.ObjPorcentajePartiDbl.ObjValorPro <>
                                        ldblPorParCli Then
                                    ReDim Preserve lstrDifsPredio(i)
                                    lstrDifsPredio(i) = lstrIdPreAgr &
                                        " las paricipaciones de los Propietarios" &
                                        " difieren!"
                                    i += 1
                                End If
                            Else
                                ReDim Preserve lstrDifsPredio(i)
                                lstrDifsPredio(i) = lstrIdPreAgr &
                                        " los propietarios difieren!"
                                i += 1
                            End If
                        Next
                    Else
                        ReDim Preserve lstrDifsPredio(i)
                        lstrDifsPredio(i) = lstrIdPreAgr &
                                " la cantidad de Propietarios difieren!"
                        i += 1
                    End If
                    If lobjPredioAgru.ObjIdTipoDestinatarioFacturaByt.ObjValorPro <>
                            lobjPre.ObjIdTipoDestinatarioFacturaByt.ObjValorPro Then
                        ReDim Preserve lstrDifsPredio(i)
                        Dim lstrDifProp = lobjPre.ObjIdPredioStr.ToString &
                                ",DestinatarioFactura"
                        lstrDifsPredio(i) = lstrDifProp
                        i += 1
                    End If
                    If lobjPredioAgru.ObjIdTipoDestinatarioFacturaByt.ObjValorPro =
                            EnuDestinatarioFacturaDef.EnuArrendatario Then
                        If lobjPredioAgru.ObjIdClienteArrendatarioDbl.ObjValorPro <>
                                lobjPre.ObjIdClienteArrendatarioDbl.ObjValorPro Then
                            ReDim Preserve lstrDifsPredio(i)
                            Dim lstrDifProp = lobjPre.ObjIdPredioStr.ToString &
                                ",Arrendatario"
                            lstrDifsPredio(i) = lstrDifProp
                            i += 1
                        End If
                    End If
                End If
            Next
        Next
        Return lstrDifsPredio
    End Function
    ''' <summary>
    ''' Devuelve  un array de datarows con la identificación de todos los predios agrupadores de la 
    ''' ubicación actual
    ''' </summary>
    Private Shared Function FdtbPrediosAgrupadores() As DataTable
        Dim lstrTabla = ClsPredio.SstrNombreTabla
        Dim lstrCamposSelect = {ClsIdPredioStr.SstrNombreCampoBd}
        Dim lstrOrden = {{ClsIdPredioStr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " & ClsIdPredioStr.SstrNombreCampoBd & " = " &
                ClsIdPredioAgrupadorStr.SstrNombreCampoBd
        Dim ldtbPreAgr = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamposSelect, lstrOrden, lstrFiltro)
        Return ldtbPreAgr
    End Function
    ''' <summary>
    ''' Devuelve el valor pasado en el argumento "adecValor" redondeado de acuerdo a la 
    ''' base de redondeo general definida en los parametros de la Copropiedad.
    ''' </summary>
    ''' <param name="adecValor">Valor a ser redondeado</param>
    ''' <returns>El valor redondeado</returns>
    ''' <remarks></remarks>
    Friend Shared Function FdecValorRedondeado(adecValor As Decimal) As Decimal
        Dim ldecValor As Decimal
        Dim ldblBaseRedondeo As Double = GobjParametros.ObjBaseRedondeoGeneralDbl.ObjValorPro
        Select Case ldblBaseRedondeo
            Case Is = 0.1
                ldecValor = Math.Round(adecValor, 1)
            Case Is = 0.2
                ldecValor = Math.Round(adecValor, 2)
            Case Else
                ldblBaseRedondeo = CType(ldblBaseRedondeo, Integer)
                ldecValor = Math.Round((adecValor / ldblBaseRedondeo), 0) * ldblBaseRedondeo
        End Select
        Return ldecValor
    End Function
    ''' <summary>
    ''' Devuelve el valor pasado en el argumento "adecValor" redondeado de acuerdo a la 
    ''' base de redondeo para intereses de mora definida en los parametros de la Copropiedad.
    ''' </summary>
    ''' <param name="adecValor">Valor a ser redondeado</param>
    ''' <returns>El valor redondeado</returns>
    ''' <remarks></remarks>
    Friend Shared Function FdecValorMoraRedondeado(adecValor As Decimal) As Decimal
        Dim ldecValor As Decimal
        Dim ldblBaseRedondeo As Double = GobjParametros.ObjBaseRedondeoIntMoraDbl.ObjValorPro
        Select Case ldblBaseRedondeo
            Case Is = 0.1
                ldecValor = Math.Round(adecValor, 1)
            Case Is = 0.2
                ldecValor = Math.Round(adecValor, 2)
            Case Else
                ldblBaseRedondeo = CType(ldblBaseRedondeo, Integer)
                ldecValor = Math.Round((adecValor / ldblBaseRedondeo), 0) * ldblBaseRedondeo
        End Select
        Return ldecValor
    End Function
    Friend Shared Function FdtbDescuentos() As DataTable
        Dim ldtbDecuentos As New DataTable
        Dim ldclOrdinal As New DataColumn("Ordinal", System.Type.GetType("System.Int16"))
        Dim ldclIdFacura As New DataColumn("NroFactura", System.Type.GetType("System.String"))
        Dim ldclIdItemFac As New DataColumn("IdItemFact", System.Type.GetType("System.String"))
        Dim ldclIdTipoDcto As New DataColumn("IdTipoDcto", System.Type.GetType("System.Byte"))
        Dim ldclTipoDcto As New DataColumn("TipoDcto", System.Type.GetType("System.String"))
        Dim ldclBase As New DataColumn("Base", System.Type.GetType("System.Decimal"))
        Dim ldclTasa As New DataColumn("Tasa", System.Type.GetType("System.Double"))
        Dim ldclValorDcto As New DataColumn("Valor", System.Type.GetType("System.Decimal"))
        ldtbDecuentos.Columns.Add(ldclOrdinal)
        ldtbDecuentos.Columns.Add(ldclIdFacura)
        ldtbDecuentos.Columns.Add(ldclIdItemFac)
        ldtbDecuentos.Columns.Add(ldclIdTipoDcto)
        ldtbDecuentos.Columns.Add(ldclTipoDcto)
        ldtbDecuentos.Columns.Add(ldclBase)
        ldtbDecuentos.Columns.Add(ldclTasa)
        ldtbDecuentos.Columns.Add(ldclValorDcto)
        Return ldtbDecuentos
    End Function
    Friend Shared Function FdtbDetalleDeuda() As DataTable
        Dim ldtbDetalleDeuda As New DataTable
        Dim ldclOrdinal As New DataColumn("Ordinal", System.Type.GetType("System.Int16"))
        Dim ldclNroFacura As New DataColumn("NroFactura", System.Type.GetType("System.String"))
        Dim ldclFechaFact As New DataColumn("FechaFact", System.Type.GetType("System.DateTime"))
        Dim ldclFechaVence As New DataColumn("FechaVence", System.Type.GetType("System.DateTime"))
        Dim ldclValorFact As New DataColumn("ValorFactura", System.Type.GetType("System.Decimal"))
        Dim ldclDebitos As New DataColumn("DeudaCapital", System.Type.GetType("System.Decimal"))
        Dim ldclCreditos As New DataColumn("DeudaIntMoras", System.Type.GetType("System.Decimal"))
        Dim ldclSaldo As New DataColumn("DeudaTotal", System.Type.GetType("System.Decimal"))
        ldtbDetalleDeuda.Columns.Add(ldclOrdinal)
        ldtbDetalleDeuda.Columns.Add(ldclNroFacura)
        ldtbDetalleDeuda.Columns.Add(ldclFechaFact)
        ldtbDetalleDeuda.Columns.Add(ldclFechaVence)
        ldtbDetalleDeuda.Columns.Add(ldclValorFact)
        ldtbDetalleDeuda.Columns.Add(ldclDebitos)
        ldtbDetalleDeuda.Columns.Add(ldclCreditos)
        ldtbDetalleDeuda.Columns.Add(ldclSaldo)
        Return ldtbDetalleDeuda
    End Function
    Friend Shared Function FdtbFacturas(adblIdCliente As Double, astrIdPredioAgru As String,
                ByRef adtmFechaDesde As Date, adtmFechaHasta As Date) As DataTable
        Dim lstrFechaDesde = ClsPanoramaDat.FstrFechaNormalizada(adtmFechaDesde)
        Dim lstrFechaHasta = ClsPanoramaDat.FstrFechaNormalizada(adtmFechaHasta)
        GobjPanDat.SControleProcesoObj(True)
        Dim lstrSaldo = ClsDebitos_FactDec.SstrNombreCampoBd & " - " &
                ClsCreditos_FactDec.SstrNombreCampoBd & " AS Saldo"
        Dim lstrDias = "IIF(" & (ClsDebitos_FactDec.SstrNombreCampoBd & " - " &
                ClsCreditos_FactDec.SstrNombreCampoBd) & " > 0, " & "DATEDIFF(CURDATE(), " &
                ClsFechaVencimientoDtm.SstrNombreCampoBd & "),0) AS Dias"
        Dim lstrCamposSelect() = {"*", lstrSaldo, lstrDias}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " &
                ClsIdCliente_FactDbl.SstrNombreCampoBd & " = " & adblIdCliente.ToString & " AND " &
                ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd & " = '" & astrIdPredioAgru & "'"
        lstrFiltro &= " AND " & ClsFechaFacturaDtm.SstrNombreCampoBd & " >= '" & lstrFechaDesde &
                "' AND " & ClsFechaFacturaDtm.SstrNombreCampoBd & " <= '" & lstrFechaHasta & "'"
        Dim lstrIndice = {{StrCampoCarpeta, "ASC"},
                {StrCampoCentroUtil, "ASC"},
                {ClsIdCliente_FactDbl.SstrNombreCampoBd, "ASC"},
                {ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd, "ASC"},
                {ClsFechaFacturaDtm.SstrNombreCampoBd, "ASC"},
                {ClsPrefijo_FactStr.SstrNombreCampoBd, "ASC"},
                {ClsIdFacturaEnt.SstrNombreCampoBd, "ASC"}}
        Dim ldtbFacturas = ClsPanorama.FdtbDataTable(ClsFactura.SstrNombreTabla, lstrCamposSelect,
                lstrIndice, lstrFiltro)
        GobjPanDat.SControleProcesoObj(False)
        Return ldtbFacturas
    End Function
    Friend Shared Function FstrIdCtaDbDscto(aenuTipoDscto As EnuTipoDescuentoDef) As String
        Dim lstrIdCtaDb = String.Empty
        With GobjParametros
            Select Case aenuTipoDscto
                Case EnuTipoDescuentoDef.EnuDsctoCapital
                    lstrIdCtaDb = GCSTRCUENTADSCTOCAP
                Case EnuTipoDescuentoDef.EnuDsctoIntMora
                    lstrIdCtaDb = GCSTRCUENTADSCTOINT
                Case EnuTipoDescuentoDef.EnuReteCree
                    lstrIdCtaDb = String.Empty
                Case EnuTipoDescuentoDef.EnuReteFuente
                    lstrIdCtaDb = .ObjIdCtaReteFuenteStr.ObjValorPro
                Case EnuTipoDescuentoDef.EnuReteIca
                    lstrIdCtaDb = .ObjIdCtaReteIcaStr.ObjValorPro
                Case EnuTipoDescuentoDef.EnuReteIva
                    lstrIdCtaDb = .ObjIdCtaReteIvaStr.ObjValorPro
                Case EnuTipoDescuentoDef.EnuDsctoPP
                    lstrIdCtaDb = .ObjIdCtaDescuentosPPStr.ObjValorPro
                Case EnuTipoDescuentoDef.EnuCancelaIva
                    lstrIdCtaDb = .ObjIdCtaImptosAsumidosStr.ObjValorPro
            End Select
        End With
        Return lstrIdCtaDb
    End Function
    Shared Function FstrDocOrigenNovedad(aenuDocOrigen As EnuTipoDocOri) As String
        Dim lstrDocOrigen = String.Empty
        With GobjParametros
            If .ObjTipoInterfazByt.ObjValorPro = EnuTipoInterfazDef.EnuPorComprobante Then
                Select Case aenuDocOrigen
                    Case EnuTipoDocOri.EnuFactura
                        lstrDocOrigen = "FAC"
                    Case EnuTipoDocOri.EnuReciboCaja
                        lstrDocOrigen = "REC"
                    Case EnuTipoDocOri.EnuNotaDb
                        lstrDocOrigen = "NDB"
                    Case EnuTipoDocOri.EnuNotaCr
                        lstrDocOrigen = "NCR"
                    Case EnuTipoDocOri.EnuNotaCon
                        lstrDocOrigen = "NCO"
                    Case EnuTipoDocOri.EnuNotaDevAnt
                        lstrDocOrigen = "NDA"
                    Case EnuTipoDocOri.EnuNotaRevCr
                        lstrDocOrigen = "NRRC"
                    Case EnuTipoDocOri.EnuNotaAjuste
                        lstrDocOrigen = "NAC"
                End Select
            ElseIf .ObjTipoInterfazByt.ObjValorPro = EnuTipoInterfazDef.EnuPorDocumento Then
                Dim lobjDoc As ClsDocumento = .ColDocumentos(aenuDocOrigen)
                lstrDocOrigen = lobjDoc.ObjTipoDocumentoStr.ObjValorPro
            End If
        End With
        Return lstrDocOrigen
    End Function
    Friend Shared Function FenuTipoNovContraria(aenuTipoNov As EnuTipoNov) As EnuTipoNov
        Dim lenuTipoNovCon = FenuTipoNovContrariaIni(aenuTipoNov)
        If lenuTipoNovCon = EnuTipoNov.None Then
            Select Case aenuTipoNov
                Case EnuTipoNov.EnuCrRetCre
                    lenuTipoNovCon = EnuTipoNov.EnuRCrRetCre
                Case EnuTipoNov.EnuRDbCap
                    lenuTipoNovCon = EnuTipoNov.EnuDbCap
                Case EnuTipoNov.EnuRDbIva
                    lenuTipoNovCon = EnuTipoNov.EnuDbIva
                Case EnuTipoNov.EnuRDbInt
                    lenuTipoNovCon = EnuTipoNov.EnuDbInt
                Case EnuTipoNov.EnuRCrPagoCap
                    lenuTipoNovCon = EnuTipoNov.EnuCrPagoCap
                Case EnuTipoNov.EnuRCrPagoInt
                    lenuTipoNovCon = EnuTipoNov.EnuCrPagoInt
                Case EnuTipoNov.EnuRCrAnApCap
                    lenuTipoNovCon = EnuTipoNov.EnuCrAnApCap
                Case EnuTipoNov.EnuRCrAnApInt
                    lenuTipoNovCon = EnuTipoNov.EnuCrAnApInt
                Case EnuTipoNov.EnuRCrDctoCap
                    lenuTipoNovCon = EnuTipoNov.EnuCrDctoCap
                Case EnuTipoNov.EnuRCrIvaGas
                    lenuTipoNovCon = EnuTipoNov.EnuCrIvaGas
                Case EnuTipoNov.EnuRCrDctoInt
                    lenuTipoNovCon = EnuTipoNov.EnuCrDctoInt
                Case EnuTipoNov.EnuRCrRetFte
                    lenuTipoNovCon = EnuTipoNov.EnuCrRetFte
                Case EnuTipoNov.EnuRCrRetIva
                    lenuTipoNovCon = EnuTipoNov.EnuCrRetIva
                Case EnuTipoNov.EnuRCrRetIca
                    lenuTipoNovCon = EnuTipoNov.EnuCrRetIca
                Case EnuTipoNov.EnuRCrRetCre
                    lenuTipoNovCon = EnuTipoNov.EnuCrRetCre
' Anticipos
                Case EnuTipoNov.EnuCrAntRec
                    lenuTipoNovCon = EnuTipoNov.EnuRCrAntRec
                Case EnuTipoNov.EnuDbAntDev
                    lenuTipoNovCon = EnuTipoNov.EnuRDbAntDev
                Case EnuTipoNov.EnuDbAntApl
                    lenuTipoNovCon = EnuTipoNov.EnuRDbAntApl
                Case EnuTipoNov.EnuRCrAntRec
                    lenuTipoNovCon = EnuTipoNov.EnuCrAntRec
                Case EnuTipoNov.EnuRDbAntDev
                    lenuTipoNovCon = EnuTipoNov.EnuDbAntDev
                Case EnuTipoNov.EnuRDbAntApl
                    lenuTipoNovCon = EnuTipoNov.EnuDbAntApl
           ' Iva intereses mora
                Case EnuTipoNov.EnuRDbIvaInt
                    lenuTipoNovCon = EnuTipoNov.EnuDbIvaInt
            End Select
        End If
        Return lenuTipoNovCon
    End Function
    ' Complemento de la funcion anterior:fenuTipoNovContraria
    Private Shared Function FenuTipoNovContrariaIni(aenuTipoNov As EnuTipoNov) As EnuTipoNov
        Dim lenuTipoNovCon As EnuTipoNov = EnuTipoNov.None
        Select Case aenuTipoNov
            Case EnuTipoNov.EnuDbCap
                lenuTipoNovCon = EnuTipoNov.EnuRDbCap
            Case EnuTipoNov.EnuDbIva
                lenuTipoNovCon = EnuTipoNov.EnuRDbIva
            Case EnuTipoNov.EnuDbIvaInt
                lenuTipoNovCon = EnuTipoNov.EnuRDbIvaInt
            Case EnuTipoNov.EnuDbInt
                lenuTipoNovCon = EnuTipoNov.EnuRDbInt
            Case EnuTipoNov.EnuCrPagoCap
                lenuTipoNovCon = EnuTipoNov.EnuRCrPagoCap
            Case EnuTipoNov.EnuCrPagoInt
                lenuTipoNovCon = EnuTipoNov.EnuRCrPagoInt
            Case EnuTipoNov.EnuCrAnApCap
                lenuTipoNovCon = EnuTipoNov.EnuRCrAnApCap
            Case EnuTipoNov.EnuCrAnApInt
                lenuTipoNovCon = EnuTipoNov.EnuRCrAnApInt
            Case EnuTipoNov.EnuCrDctoCap
                lenuTipoNovCon = EnuTipoNov.EnuRCrDctoCap
            Case EnuTipoNov.EnuCrDctoInt
                lenuTipoNovCon = EnuTipoNov.EnuRCrDctoInt
            Case EnuTipoNov.EnuCrRetFte
                lenuTipoNovCon = EnuTipoNov.EnuRCrRetFte
            Case EnuTipoNov.EnuCrRetIva
                lenuTipoNovCon = EnuTipoNov.EnuRCrRetIva
            Case EnuTipoNov.EnuCrRetIca
                lenuTipoNovCon = EnuTipoNov.EnuRCrRetIca
            Case EnuTipoNov.EnuCrIvaGas
                lenuTipoNovCon = EnuTipoNov.EnuRCrIvaGas
        End Select
        Return lenuTipoNovCon
    End Function
    ''' <summary>
    ''' Devuelve la lista de los Predios Agrupadores correspondientes a los predios
    ''' pasados en el argumento "astrIdPredios"
    ''' </summary>
    ''' <param name="astrIdPredios">Lista de los predios cuyo predio agrupador formara
    ''' parte de la lista devuelta por la función.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Shared Function FstrIdPrediosAgr(astrIdPredios As String()) As ArrayList
        Dim lstrIdPrediosAgr As New ArrayList
        Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.EnuUnico)
        For Each lstrIdPred As String In astrIdPredios
            If String.IsNullOrEmpty(lstrIdPred) Then
                If Not lstrIdPrediosAgr.Contains(lstrIdPred) Then
                    lstrIdPrediosAgr.Add(lstrIdPred)
                End If
            Else
                lobjPredio.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrIdPred})
                If Not lstrIdPrediosAgr.Contains(lobjPredio.ObjIdPredioAgrupadorStr.ObjValorPro) Then
                    lstrIdPrediosAgr.Add(lobjPredio.ObjIdPredioAgrupadorStr.ObjValorPro)
                End If
            End If
        Next
        lstrIdPrediosAgr.Sort()
        Return lstrIdPrediosAgr
    End Function
    ''' <summary>
    ''' Indica si todos los predios agrupadores tienen un valor para el servicio de identificación
    ''' </summary>
    Friend Shared Function FblnTodosPrediosAgruConValorId() As Boolean
        Dim lstrFiltro = StrFiltroUbicacion & " AND " & ClsIdPredioStr.SstrNombreCampoBd +
                " = " & ClsIdPredioAgrupadorStr.SstrNombreCampoBd & " AND " &
                ClsValorServicioIdDec.SstrNombreCampoBd & " = 0"
        Dim ldtbPrediosConValorId = ClsPanorama.FdtbDataTable(ClsPredio.SstrNombreTabla,
                {ClsIdPredioStr.SstrNombreCampoBd}, {{"", ""}}, lstrFiltro)
        Return ldtbPrediosConValorId.Rows.Count = 0
    End Function
    Friend Shared Function FblnDocReversable(astrPrefDoc As String, aentIdDoc As Integer,
            aenuDocReversado As EnuDocReversado, ByRef astrMens As String) As Boolean
        Dim lblnRever As Boolean
        Dim lstrTabla = ClsNotaReversionCr.SstrNombreTabla
        Dim lstrCamposSelect = {ClsIdNotaReversaCrEnt.SstrNombreCampoBd}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " &
                ClsPrefijoDoc_NotaReversaCrStr.SstrNombreCampoBd & " = '" & astrPrefDoc &
                "' AND " & ClsIdDoc_NotaReversaCrEnt.SstrNombreCampoBd & " = " &
                aentIdDoc.ToString & " AND " & ClsTipoDocReversadoByt.SstrNombreCampoBd & " = " &
                aenuDocReversado
        Dim ldtbNotaRevCr = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamposSelect, {{"", ""}},
                lstrFiltro)
        lblnRever = (ldtbNotaRevCr.Rows.Count = 0)
        If lblnRever Then
            Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, astrPrefDoc, aentIdDoc}
            If aenuDocReversado = EnuDocReversado.EnuReciboC Then
                Dim lobjRec As New ClsReciboCaja()
                lobjRec.SAbra(lobjValorLlave)
                If Not IsNothing(lobjRec.ObjAnticipo) Then
                    lblnRever = Not lobjRec.ObjAnticipo.FblnAntReintegrado
                    If Not lblnRever Then
                        astrMens = "El Recibo de Caja no puede ser reversado; " &
                                "tiene un Anticipo reintegrado!"
                    End If
                End If
            ElseIf aenuDocReversado = EnuDocReversado.EnuNotaCr Then
                Dim lobjNcr As New ClsNotaCr()
                lobjNcr.SAbra(lobjValorLlave)
                If lobjNcr.ObjIdTipoNotaCrByt.ObjValorPro = EnuTipoNotaCrDef.EnuRetenciones Then
                    astrMens = "La reversión de esta Nota Crédito no se " &
                            "registrará en la API de Facturación Electrónica!"
                End If
            End If
        Else
            astrMens = "El Documento seleccionado ya fue reversado!"
        End If
        Return lblnRever
    End Function
    Friend Shared Function FblnPuedeCrear(aenuTipoDoc As EnuTipoDocOri, ablnFactManual As Boolean,
            ByRef astrMens As String) As Boolean
        Dim lblnPuede As Boolean
        If GobjParametros.BlnEFacAutorizado Then
            lblnPuede = GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro
            If Not lblnPuede Then
                astrMens = "En EFac se requiere que la Opción: " & Chr(34) &
                        "Documentos deben ser expedidos con Fecha de Hoy" & Chr(34) & ", este activa!"
                Return lblnPuede
            End If
        End If
        Dim lobjCentroUtil As ClsCentroUtilidad = GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual
        lblnPuede = lobjCentroUtil.ObjEstadoContratoByt.ObjValorPro <>
                    EnuEstadoContrato.EnuSuspendido
        If Not lblnPuede Then
            astrMens = "No es posible la creación. El contrtrato está en estado suspendido!"
        End If
        If lblnPuede Then
            If aenuTipoDoc = EnuTipoDocOri.EnuFactura Then
                lblnPuede = FblnResDianOK(astrMens)
                If lblnPuede Then
                    If Not ablnFactManual Then
                        If BlnPreFacturando Then
                            lblnPuede = FblnHayItemsPorFacturar()
                            If Not lblnPuede Then
                                astrMens = "No hay Servicios para ser Pre-Facturados!"
                            Else
                                lblnPuede = Not FblnHayPrefacturas()
                                If Not lblnPuede Then
                                    astrMens = "Hay Pre-Facturas generadas!"
                                End If
                            End If
                        Else
                            lblnPuede = FblnHayPrefacturas()
                            If Not lblnPuede Then
                                astrMens = "No hay Pre-Facturas generadas!"
                            End If
                        End If
                    End If
                End If
            Else
                lblnPuede = Not FblnHayPrefacturas()
                If Not lblnPuede Then
                    astrMens = " si hay procesos de Facturación pendientes de llevarse a cabo!"
                Else
                    If GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro Then
                        lblnPuede = Not ClsOrionCop.FblnHacerCierreMes
                        If Not lblnPuede Then
                            astrMens = " si no se ha cerradp el Periódo actual!"
                        End If
                    Else
                        lblnPuede = Not (FblnHayItemsPorFacturar())
                        If Not lblnPuede Then
                            astrMens = " si hay procesos de Facturación pendientes de " &
                                    "llevarse a cabo!"
                        End If
                    End If
                End If
                If Not lblnPuede Then
                    Dim lstrDoc = String.Empty
                    If aenuTipoDoc = EnuTipoDocOri.EnuNotaCr Then
                        lstrDoc = "Notas Crédito"
                    ElseIf aenuTipoDoc = EnuTipoDocOri.EnuReciboCaja Then
                        lstrDoc = "Recibos de Caja"
                    ElseIf aenuTipoDoc = EnuTipoDocOri.EnuNotaRevCr Then
                        lstrDoc = "Nota Reversión Recibo de Caja"
                    End If
                    astrMens = "No está permitido generar " & lstrDoc & astrMens
                End If
            End If
        End If
        Return lblnPuede
    End Function
    Private Shared Function FblnResDianOK(ByRef astrMens As String) As Boolean
        Dim lblnPuede = True
        If Not String.IsNullOrEmpty(GobjParametros.ObjNumeroResolFacturaStr.ToString()) Then
            Dim lenuEstadoResDian = GobjParametros.FEnuEstadoResDian
            If lenuEstadoResDian = EnuEstadoResDian.EnuSinResVigente Then
                lblnPuede = False
                astrMens = "No es posible generar facturas! No hay Resolución vigente de la DIAN!"
            ElseIf lenuEstadoResDian = EnuEstadoResDian.EnuVencida Then
                lblnPuede = False
                astrMens = "No es posible generar facturas! La Resolución de la DIAN está vencida!"
            ElseIf lenuEstadoResDian = EnuEstadoResDian.EnuNumAgotada Then
                lblnPuede = False
                astrMens = "No es posible generar facturas! La númeración aprobada por la DIAN " &
                        "está agotada!!"
            End If
        End If
        Return lblnPuede
    End Function
    Friend Shared Function FentCantidadPredios() As Integer
        Dim lstrFiltro = StrFiltroUbicacion
        Dim ldtbPredios = ClsPanorama.FdtbDataTable(ClsPredio.SstrNombreTabla,
                {"COUNT(" & ClsIdPredioStr.SstrNombreCampoBd & ")"}, {{"", ""}},
                lstrFiltro)
        Return ldtbPredios.Rows(0)(0)
    End Function
    ''' <summary>
    ''' Indica si las cuotas de Administración del año actual ya fueron calculadas por primera vez.
    ''' </summary>
    ''' <returns></returns>
    Friend Shared Function FblnEstanCalcuCuotasAdmin(ashrIdAno As Short) As Boolean
        Dim lblnEstan = False
        If GobjParametros.ObjAnoActual IsNot Nothing Then
            Dim lstrTabla = ClsItemProgramaFact.SstrNombreTabla
            Dim lstrCamSelect = {"COUNT(" & ClsIdPredioStr.SstrNombreCampoBd & ")"}
            Dim lstrFiltro = StrFiltroUbicacion & " AND " &
                    ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " = " & ashrIdAno
            Dim ldtbRes = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamSelect, {{"", ""}}, lstrFiltro)
            lblnEstan = ldtbRes.Rows(0)(0) > 0
        End If
        Return lblnEstan
    End Function
    ''' <summary>
    ''' indica si en el año actual hay por lo menos un período facturado
    ''' </summary>
    ''' <returns></returns>
    Friend Shared Function FblnHayPeriFacturados() As Boolean
        Dim lblnHay = False, lentPerPorFacturar As Integer
        Dim lstrPerIni As String, lentCanPer As Integer, ldecVlrPer As Decimal, ldecSaldo As Decimal
        Dim lstrIdAno = GobjParametros.ObjAnoActual.ObjIdAnoShr.ToString()
        Dim lstrTabla = ClsItemProgramaFact.SstrNombreTabla
        Dim lstrCamSel = {ClsPeriodoIni_ItemProgStr.SstrNombreCampoBd,
                ClsCantidadPeriodosShr.SstrNombreCampoBd,
                ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd,
                ClsSaldo_ItemProgramaFactDec.SstrNombreCampoBd}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " &
                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " = " & lstrIdAno & " AND " &
                ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd & " = 1 AND " &
                ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd & " > 0 "
        Dim ldtbItemsProFac = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamSel, {{"", ""}}, lstrFiltro)
        For Each ldrwItem As DataRow In ldtbItemsProFac.Rows
            lstrPerIni = ClsPanorama.FobjValorCampo(ldrwItem(
                    ClsPeriodoIni_ItemProgStr.SstrNombreCampoBd), EnuTipoValor.EnuString)
            lentCanPer = ClsPanorama.FobjValorCampo(ldrwItem(
                    ClsCantidadPeriodosShr.SstrNombreCampoBd), EnuTipoValor.EnuInteger)
            ldecVlrPer = ClsPanorama.FobjValorCampo(ldrwItem(
                    ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd), EnuTipoValor.EnuDecimal)
            ldecSaldo = ClsPanorama.FobjValorCampo(ldrwItem(
                    ClsSaldo_ItemProgramaFactDec.SstrNombreCampoBd), EnuTipoValor.EnuDecimal)
            lentPerPorFacturar = ldecSaldo / ldecVlrPer
            lblnHay = lentPerPorFacturar < lentCanPer
            If lblnHay Then Exit For
        Next
        Return lblnHay
    End Function
    ''' <summary>
    ''' Devuelve un arreglo de datarows que contiene un campo con los prefijos usados en el 
    ''' documento pasado en el parámetro
    ''' </summary>
    Friend Shared Function FdrwPrefDoc(aenuTipoDoc As EnuTipoDocOri) As DataRow()
        Dim lstrNomTabla As String
        Dim lstrNomCampo As String
        Select Case aenuTipoDoc
            Case EnuTipoDocOri.EnuFactura
                lstrNomTabla = ClsFactura.SstrNombreTabla
                lstrNomCampo = ClsPrefijo_FactStr.SstrNombreCampoBd
            Case EnuTipoDocOri.EnuNotaAjuste
                lstrNomTabla = ClsNotaAjusteCuotaAdmin.SstrNombreTabla
                lstrNomCampo = ClsPrefijo_NotaAjusteStr.SstrNombreCampoBd
            Case EnuTipoDocOri.EnuNotaCon
                lstrNomTabla = ClsNotaCon.SstrNombreTabla
                lstrNomCampo = ClsPrefijo_NotaConStr.SstrNombreCampoBd
            Case EnuTipoDocOri.EnuNotaCr
                lstrNomTabla = ClsNotaCr.SstrNombreTabla
                lstrNomCampo = ClsPrefijo_NotaCrStr.SstrNombreCampoBd
            Case EnuTipoDocOri.EnuNotaDb
                lstrNomTabla = ClsNotaDb.SstrNombreTabla
                lstrNomCampo = ClsPrefijo_NotaDbStr.SstrNombreCampoBd
            Case EnuTipoDocOri.EnuNotaDevAnt
                lstrNomTabla = ClsNotaDevAnt.SstrNombreTabla
                lstrNomCampo = ClsPrefijo_NotaDevAntStr.SstrNombreCampoBd
            Case EnuTipoDocOri.EnuNotaRevCr
                lstrNomTabla = ClsNotaReversionCr.SstrNombreTabla
                lstrNomCampo = ClsPrefijo_NotaReversaCrStr.SstrNombreCampoBd
            Case EnuTipoDocOri.EnuReciboCaja
                lstrNomTabla = ClsReciboCaja.SstrNombreTabla
                lstrNomCampo = ClsPrefijo_RecStr.SstrNombreCampoBd
            Case Else
                Throw New ErrorInesperadoPanLException("Tipo de Documento invalido!")
        End Select
        Dim lstrCampSele As String() = {"DISTINCT(" & lstrNomCampo & ") AS Dato"}
        Dim lstrOrden As String(,) = {{lstrNomCampo, "ASC"}}
        Dim lstrFiltro = StrFiltroUbicacion
        Dim ldtbPref = ClsPanorama.FdtbDataTable(lstrNomTabla, lstrCampSele, lstrOrden, lstrFiltro)
        Return ldtbPref.Select()
    End Function
    Friend Shared Function FdblTasaDscto(astrNroFact As String, ashrIdItemFac As Short,
            aenuTipoDscto As EnuTipoDescuentoDef, adecValorDscto As Decimal,
            ByRef adecBaseDscto As Decimal) As Double
        Dim lobjFac As New ClsFactura()
        Dim lstrPref = ClsPanorama.FstrPrefijoDcto(astrNroFact)
        Dim lentIdFac = ClsPanorama.FentIdDcto(astrNroFact)
        Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentIdFac}
        lobjFac.SAbra(lobjValorLlave)
        Dim lobjItemFac As ClsItemFactura = lobjFac.ColItemsFactura(ashrIdItemFac.ToString)
        Select Case aenuTipoDscto
            Case EnuTipoDescuentoDef.EnuDsctoCapital
                adecBaseDscto = lobjItemFac.FdecDeudaServicioTotal
            Case EnuTipoDescuentoDef.EnuDsctoIntMora
                adecBaseDscto = lobjItemFac.FdecDeudaIntMora
            Case EnuTipoDescuentoDef.EnuDsctoPP
                adecBaseDscto = lobjItemFac.FdecValorServicio
            Case EnuTipoDescuentoDef.EnuReteFuente
                adecBaseDscto = lobjItemFac.FdecValorServicio
            Case EnuTipoDescuentoDef.EnuReteIca
                adecBaseDscto = lobjItemFac.FdecValorServicio
            Case EnuTipoDescuentoDef.EnuReteIva
                adecBaseDscto = lobjItemFac.FdecIvaServicio
            Case EnuTipoDescuentoDef.EnuCancelaIva
                adecBaseDscto = lobjItemFac.FdecDeudaIva
        End Select
        Dim ldblTasa As Double = adecValorDscto / adecBaseDscto
        Return ldblTasa
    End Function
    Friend Shared Function FdblTasaDscto(aobjItemFact As ClsItemFactura,
            aenuTipoDscto As EnuTipoDescuentoDef, adecValorDscto As Decimal,
            ByRef adecBaseDscto As Decimal) As Double
        Select Case aenuTipoDscto
            Case EnuTipoDescuentoDef.EnuDsctoCapital
                adecBaseDscto = aobjItemFact.FdecDeudaServicioTotal
            Case EnuTipoDescuentoDef.EnuDsctoIntMora
                adecBaseDscto = aobjItemFact.FdecDeudaIntMora
            Case EnuTipoDescuentoDef.EnuDsctoPP
                adecBaseDscto = aobjItemFact.FdecValorServicio
            Case EnuTipoDescuentoDef.EnuReteFuente
                adecBaseDscto = aobjItemFact.FdecValorServicio
            Case EnuTipoDescuentoDef.EnuReteIca
                adecBaseDscto = aobjItemFact.FdecValorServicio
            Case EnuTipoDescuentoDef.EnuReteIva
                adecBaseDscto = aobjItemFact.FdecIvaServicio
            Case EnuTipoDescuentoDef.EnuCancelaIva
                adecBaseDscto = aobjItemFact.FdecDeudaIva
        End Select
        Dim ldblTasa As Double = adecValorDscto / adecBaseDscto
        Return ldblTasa
    End Function
    ''' <summary>
    ''' Devuelve el valor anual del servicio definido por los argumentos, es decir lo que realmente 
    ''' está en la programación de facturación
    ''' </summary>
    Friend Shared Function FdecValorTotalCalculadoServicio(ashrIdAno As Short,
            ashrIdServicio As Short) As Decimal
        Dim ldecValor = 0D
        Dim lstrCamposSelect As String() = {"SUM(" &
                ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd & " * " &
                ClsCantidadPeriodosShr.SstrNombreCampoBd & ") AS Valor"}
        Dim lstrfiltro = StrFiltroUbicacion & " And " &
                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " = " & ashrIdAno &
                " And " & ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd & " = " &
                ashrIdServicio
        Dim ldtbValorTotal As DataTable = ClsPanorama.FdtbDataTable(ClsItemProgramaFact.SstrNombreTabla,
                lstrCamposSelect, {{"", ""}}, lstrfiltro)
        If ldtbValorTotal.Rows.Count > 0 Then
            Dim ldrwVlrTot As DataRow = ldtbValorTotal.Rows(0)
            ldecValor = ClsPanorama.FobjValorCampo(ldrwVlrTot(0), EnuTipoValor.enuDecimal)
        End If
        Return ldecValor
    End Function
    Friend Shared Sub SReasigneAliasCont(astrAliasContActual As String, astrAliasConNuevo As String)
        Dim lstrTabla = ClsNovedad.SstrNombreTabla
        Dim lcolCampCambio As New Collection From {
            ClsAliasCont_NovStr.SstrNombreCampoBd
        }
        Dim lcolDatosNuevos As New Collection From {
            {astrAliasConNuevo, ClsAliasCont_NovStr.SstrNombreCampoBd}
        }
        Dim lcolCamposRef As New Collection From {
            StrCampoCarpeta,
            StrCampoCentroUtil,
            ClsAliasCont_NovStr.SstrNombreCampoBd
        }
        Dim lcolDatosRef As New Collection From {
                {GshrIdCarpeta, StrCampoCarpeta},
                {GshrIdCentroUtil, StrCampoCentroUtil},
                {astrAliasContActual, ClsAliasCont_NovStr.SstrNombreCampoBd}
        }
        GobjPanDat.SActualiceRegistro(lstrTabla, lcolCampCambio, lcolDatosNuevos,
                lcolCamposRef, lcolDatosRef)
    End Sub
    Friend Shared Function FdtbCoeficientesPropPropietarios() As DataTable
        Dim lstrExpSql = "SELECT PRO." & ClsIdCliente_PropDbl.SstrNombreCampoBd & ", " &
                FstrNombres() & FstrApellidos() & FstrRazSocial() & FstrTel() & FstrCorreo() &
                FstrPredio() & FstrCP() & "FROM " & ClsPredio.SstrNombreTabla &
                " AS PRE INNER JOIN " & ClsPropietario.SstrNombreTabla & " AS PRO On PRE." &
                StrCampoCarpeta & " = PRO." & StrCampoCarpeta &
                " AND PRE." & ClsIdPredioStr.SstrNombreCampoBd & " = PRO." &
                ClsIdPredio_PropStr.SstrNombreCampoBd & " INNER JOIN " & ClsTercero.SstrNombreTabla &
                " AS TR On PRO." & ClsIdCliente_PropDbl.SstrNombreCampoBd & " = TR." &
                ClsIdTerceroDbl.SstrNombreCampoBd & " WHERE PRE." &
                StrCampoCarpeta & " = " & GshrIdCarpeta & " AND PRE." &
                StrCampoCentroUtil & " = " & GshrIdCentroUtil & " AND " &
                ClsCoeficientePropiedadDec.SstrNombreCampoBd & " > 0 GROUP BY PRE." &
                ClsIdPredioAgrupadorStr.SstrNombreCampoBd
        Dim ldtbCPP = ClsPanorama.FdtbDataTable(lstrExpSql)
        Return ldtbCPP
    End Function
    Private Shared Function FstrNombres() As String
        Dim lstrNombres = "If(TR." & ClsNombreSegundoStr.SstrNombreCampoBd & " <> '', CONCAT(TR." &
                ClsNombrePrimeroStr.SstrNombreCampoBd & ", ' ', TR." &
                ClsNombreSegundoStr.SstrNombreCampoBd & "), TR." &
                ClsNombrePrimeroStr.SstrNombreCampoBd & ") AS Nombres, "
        Return lstrNombres
    End Function
    Private Shared Function FstrApellidos() As String
        Dim lstrApellidos = "IF(TR." & ClsApellidoSegundoStr.SstrNombreCampoBd & " <> '',CONCAT(TR." &
                ClsApellidoPrimeroStr.SstrNombreCampoBd & ",' ', TR." &
                ClsApellidoSegundoStr.SstrNombreCampoBd & "), TR." &
                ClsApellidoPrimeroStr.SstrNombreCampoBd & " ) As Apellidos, "
        Return lstrApellidos
    End Function
    Private Shared Function FstrRazSocial() As String
        Return "TR." & ClsRazonSocialStr.SstrNombreCampoBd & " AS 'RazonSocial', "
    End Function
    Private Shared Function FstrTel() As String
        Dim lstrTel = "IF(TR." & ClsCelularStr.SstrNombreCampoBd & " = '',IF(TR." &
                ClsTelefonoUnoStr.SstrNombreCampoBd & "<> '',TR." &
                ClsTelefonoUnoStr.SstrNombreCampoBd & ", ''), TR." &
                ClsCelularStr.SstrNombreCampoBd & ") AS Telefono, "
        Return lstrTel
    End Function
    Private Shared Function FstrCorreo() As String
        Dim lstrCorreo = "IF(TR." & ClsEmailStr.SstrNombreCampoBd & " = '', IF(PRE." &
                ClsEmailAdiStr.SstrNombreCampoBd & " <> '', PRE." &
                ClsEmailAdiStr.SstrNombreCampoBd & ",''), TR." & ClsEmailStr.SstrNombreCampoBd &
                ") AS Correo, "
        Return lstrCorreo
    End Function
    Private Shared Function FstrPredio() As String
        Return "PRE." & ClsIdPredioAgrupadorStr.SstrNombreCampoBd & " AS Predio, "
    End Function
    Private Shared Function FstrCP() As String
        Dim lstrCoefProp = "SUM(PRE." & ClsCoeficientePropiedadDec.SstrNombreCampoBd & " * " &
                ClsPorcentajePartiDbl.SstrNombreCampoBd & ") AS 'TotalCoeficientes' "
        Return lstrCoefProp
    End Function
#End Region
#Region "Calculos Predios"
    ''' <summary>
    ''' Calcula el total de la base de calculo para el coeficiente de propiedad y calcula el
    ''' coeficiente de propiedad de cada uno de los predios. Ademas calcula la Base de Participación 
    ''' de los Sectores que conforman la Copropiedad. Actualiza los datos en la BD.
    ''' </summary>
    ''' <remarks></remarks>
    Friend Shared Sub SActualiceTotalAreaCopr()
        Dim ldecTotalBaseParticipacion = 0D, ldecTotalAreaPrivada = 0D, lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            GobjPanDat.SInicialiceTransaccion()
            Dim lstrTabla = ClsPredio.SstrNombreTabla
            Dim lstrCampSel = "SUM(" & ClsAreaPredioDec.SstrNombreCampoBd & ") AS TOTA ," &
                    "SUM(" & ClsAreaPredioDec.SstrNombreCampoBd & " * " &
                    ClsFactorPonderaCPDbl.SstrNombreCampoBd & ") As TotalAreaPond"
            Dim lstrExpSql = "SELECT " & lstrCampSel & " FROM " & lstrTabla &
                    " WHERE " & StrFiltroUbicacion
            Dim ldtbCalculo = ClsPanorama.FdtbDataTable(lstrExpSql)
            If ldtbCalculo.Rows.Count > 0 Then
                Dim ldrwCalculo As DataRow = ldtbCalculo.Rows(0)
                ldecTotalAreaPrivada = Math.Round(ClsPanorama.FobjValorCampo(ldrwCalculo(0),
                        EnuTipoValor.enuDecimal), 4)
                ldecTotalBaseParticipacion = Math.Round(ClsPanorama.FobjValorCampo(ldrwCalculo(1),
                        EnuTipoValor.enuDecimal), 4)
            End If
            If ldecTotalBaseParticipacion <> GobjParametros.ObjTotalAreaPondDec.ObjValorPro OrElse
                    ldecTotalAreaPrivada <> GobjParametros.ObjTotalAreaCopropDec.ObjValorPro Then
                GobjParametros.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                GobjParametros.ObjTotalAreaCopropDec.ObjValorPro = ldecTotalAreaPrivada
                GobjParametros.ObjTotalAreaPondDec.ObjValorPro = ldecTotalBaseParticipacion
                GobjParametros.SActualice(False)
            End If
            SActualiceCoeficientesProp()
            lblnNoHayError = True
        Catch ex As ProveedorBdPanException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                GobjPanDat.SConfirmeTransaccion()
                GobjPanDat.SControleProcesoObj(False)
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    Private Shared Sub SActualiceCoeficientesProp()
        Dim lobjOrionCop = New ClsOrionCop(GCOBJREGISTRO, False)
        Dim lobjPredio As ClsPredio = FobjNuevoPredio(EnuModoInstanciaObjDef.enuNavegable)
        lobjPredio.SVayaAlPrimero()
        Dim ldecCoeficientePro = 0D
        Do While lobjPredio.BlnExiste
            With lobjPredio
                ldecCoeficientePro = .ObjAreaPredioDec.ObjValorPro *
                        .ObjFactorPonderaCPDbl.ObjValorPro /
                        GobjParametros.ObjTotalAreaPondDec.ObjValorPro
                Select Case GobjParametros.ObjBaseRedondeoCPByt.ObjValorPro
                    Case 1
                        .ObjCoeficientePropiedadDec.ObjValorPro = ldecCoeficientePro
                    Case Else
                        ldecCoeficientePro = Math.Round(ldecCoeficientePro,
                                GobjParametros.ObjBaseRedondeoCPByt.ObjValorPro + 2)
                        .ObjCoeficientePropiedadDec.ObjValorPro = ldecCoeficientePro
                End Select
                .SModifique()
                .SActualice(False)
            End With
            lobjPredio.SVayaAlSiguiente()
        Loop
    End Sub
#End Region
#Region "Integridad"
    Friend Function SCompruebeIntegridad(ByRef ablnHayError As Boolean,
                adtmfechaIni As Date, adtmfechaFin As Date,
                aenuDocInt As EnuDocsIntegridad) As String
        MobjArgumentoEventoPan = Nothing
        Dim lblnSeguir As Boolean = True
        Dim lstrFechaDesde = "'" & ClsPanoramaDat.FstrFechaNormalizada(adtmfechaIni.ToString) & "'"
        Dim lstrFechaHasta = "'" & ClsPanoramaDat.FstrFechaNormalizada(adtmfechaFin.ToString) & "'"
        GobjPanDat.SControleProcesoObj(True)
        MblnHayerror = False
        SInicialiceArchivos()
        BlnCorriendoIntegridad = True
        ObjArgumentoEventoPan.BlnCancele = False
        If aenuDocInt And EnuDocsIntegridad.EnuFac Then
            SCompruebeFacturas(lstrFechaDesde, lstrFechaHasta)
            If ObjArgumentoEventoPan.BlnCancele Then
                BlnCorriendoIntegridad = False
                lblnSeguir = False
            End If
        End If
        If lblnSeguir Then
            If aenuDocInt And EnuDocsIntegridad.EnuRC Then
                SCompruebeRecs(lstrFechaDesde, lstrFechaHasta)
                If ObjArgumentoEventoPan.BlnCancele Then
                    BlnCorriendoIntegridad = False
                    lblnSeguir = False
                End If
            End If
        End If
        If lblnSeguir Then
            If aenuDocInt And EnuDocsIntegridad.EnuNcr Then
                SCompruebeNotasCr(lstrFechaDesde, lstrFechaHasta)
                If ObjArgumentoEventoPan.BlnCancele Then
                    BlnCorriendoIntegridad = False
                    lblnSeguir = False
                End If
            End If
        End If
        If lblnSeguir Then
            If aenuDocInt And EnuDocsIntegridad.EnuNdb Then
                SCompruebeNotasDb(lstrFechaDesde, lstrFechaHasta)
                If ObjArgumentoEventoPan.BlnCancele Then
                    BlnCorriendoIntegridad = False
                    lblnSeguir = False
                End If
            End If
        End If
        If lblnSeguir Then
            If aenuDocInt And EnuDocsIntegridad.EnuNco Then
                SCompruebeNotasCon(lstrFechaDesde, lstrFechaHasta)
                If ObjArgumentoEventoPan.BlnCancele Then
                    BlnCorriendoIntegridad = False
                    lblnSeguir = False
                End If
            End If
        End If
        If lblnSeguir Then
            If aenuDocInt And EnuDocsIntegridad.EnuAnt Then
                SCompruebeAnticipos(lstrFechaDesde, lstrFechaHasta)
                If ObjArgumentoEventoPan.BlnCancele Then
                    BlnCorriendoIntegridad = False
                    lblnSeguir = False
                End If
            End If
        End If
        If lblnSeguir Then
            If aenuDocInt And EnuDocsIntegridad.EnuNrrc Then
                SCompruebeNotasRCr(lstrFechaDesde, lstrFechaHasta)
                If ObjArgumentoEventoPan.BlnCancele Then
                    BlnCorriendoIntegridad = False
                    lblnSeguir = False
                End If
            End If
        End If
        If lblnSeguir Then
            If aenuDocInt And EnuDocsIntegridad.EnuEstadoCta Then
                SCompruebeEstadoCtas(lstrFechaDesde, lstrFechaHasta)
                If ObjArgumentoEventoPan.BlnCancele Then
                    BlnCorriendoIntegridad = False
                End If
            End If
            GobjPanDat.SControleProcesoObj(False)
            If ObjArgumentoEventoPan.BlnCancele Then
                BlnCorriendoIntegridad = False
            End If
        End If
        BlnCorriendoIntegridad = False
        SAdicioneFinIntegridad("Fin del proceso que verifica la integridad de los datos.")
        ablnHayError = MblnHayerror
        Return MstrArchivoIntegridad
    End Function
    Private Sub SCompruebeFacturas(astrFechaIni As String, astrFechaFin As String)
        Dim lstrCamposSelect = {"DISTINCT(" & ClsIdFactura_NovEnt.SstrNombreCampoBd & ")",
                ClsPrefijoFact_NovStr.SstrNombreCampoBd}
        Dim lstrIndice(,) = {{ClsPrefijoFact_NovStr.SstrNombreCampoBd, "ASC"},
                {ClsIdFactura_NovEnt.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " & ClsFechaNovedadDtm.SstrNombreCampoBd &
                " >= " & astrFechaIni & " AND " & ClsFechaNovedadDtm.SstrNombreCampoBd &
                " <= " & astrFechaFin
        Dim ldtbIdFacturas = ClsPanorama.FdtbDataTable(ClsNovedad.SstrNombreTabla, lstrCamposSelect,
                lstrIndice, lstrFiltro)
        ObjArgumentoEventoPan.EnuProceso = EnuProcesoDef.enuIntegrFac
        ObjArgumentoEventoPan.DblCantAProcesar = ldtbIdFacturas.Rows.Count
        ObjArgumentoEventoPan.DblCantProcesada = 0
        RaiseEvent EvnInicio(Me, ObjArgumentoEventoPan)
        Dim i = 0
        Dim lobjFactura As New ClsFactura()
        For Each ldrwIdFact As DataRow In ldtbIdFacturas.Rows
            i += 1
            Dim lstrPrefijo As String = ClsPanorama.FobjValorCampo(ldrwIdFact(
                    ClsPrefijoFact_NovStr.SstrNombreCampoBd), EnuTipoValor.enuString)
            Dim lentIdFac As Integer = ClsPanorama.FobjValorCampo(ldrwIdFact(
                    ClsIdFactura_NovEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
            SCompruebeFactura(lobjFactura, lstrPrefijo, lentIdFac)
            ObjArgumentoEventoPan.DblCantProcesada = i
            RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
            If ObjArgumentoEventoPan.BlnCancele Then
                Exit For
            End If
        Next
    End Sub
    Private Sub SCompruebeFactura(aobjFactura As ClsFactura, astrPrefijo As String,
    aentIdFactura As Integer)
        Dim lobjValorLlave() As Object = {GshrIdCarpeta, GshrIdCentroUtil, astrPrefijo, aentIdFactura}
        Dim lblnNoHayError = False
        Try
            With aobjFactura
                Dim lstrLinea = String.Empty
                .SAbra(lobjValorLlave)
                If Not .BlnExiste Then
                    Throw New ErrorInesperadoPanLException("No existe Factura con Id de Tabla " &
                            astrPrefijo & "-" & aentIdFactura.ToString &
                            "  .No se puede comprobar Integridad!")
                End If
                Dim ldecSaldoFac As Decimal = .DecDeuda
                If ldecSaldoFac < 0 Then
                    lstrLinea = "Factura " & .ObjIdFacturaEnt.ToString & " con saldo negativo!"
                    SAdicioneErrorIntegridad(lstrLinea)
                Else
                    Dim lcolItemsFac = .ColItemsFactura
                    Dim ldecSaldoItems = 0D
                    ' Verificar Integridad de Novedades
                    Dim lshrIdItemFac = 0S
                    For Each lobjItemfac As ClsItemFactura In lcolItemsFac
                        lshrIdItemFac += 1
                        If lobjItemfac.ObjIdItemFacturaShr.ObjValorPro <> lshrIdItemFac Then
                            lstrLinea = "Factura " & .ObjIdFacturaEnt.ToString &
                                    " con numeracion de items errada!"
                            SAdicioneErrorIntegridad(lstrLinea)
                            Exit For
                        End If
                        ldecSaldoItems += lobjItemfac.DecDeuda
                        SVerifiqueIntegridadNovedades(lobjItemfac)
                    Next
                    If String.IsNullOrEmpty(lstrLinea) Then
                        If ldecSaldoFac <> ldecSaldoItems Then
                            lstrLinea = "En factura " & .ObjIdFacturaEnt.ToString &
                                    " difiere encabezado con items"
                            SAdicioneErrorIntegridad(lstrLinea)
                        End If
                    End If
                End If
            End With
            lblnNoHayError = True
        Catch ex As PanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Sub
    Private Sub SVerifiqueIntegridadNovedades(aobjItemfact As ClsItemFactura)
        Dim ldecDbItem As Decimal = aobjItemfact.ObjDebitos_ItemFactDec.ObjValorPro
        Dim ldecCrItem As Decimal = aobjItemfact.ObjCreditos_ItemFactDec.ObjValorPro
        Dim ldecDbNovs = 0D
        Dim ldecCrNovs = 0D
        Dim ldecVlrNov As Decimal
        Dim lcolNovedades = aobjItemfact.ColNovedades
        Dim lenuIdTipoNov As EnuTipoNov
        Dim lstrLinea As String
        Dim lobjFac As ClsFactura = aobjItemfact.ObjPadre
        Dim lstrIdpreAgrFac As String = lobjFac.ObjIdPredioAgrupador_FacStr.ObjValorPro
        If lcolNovedades.Count > 0 Then
            For Each lobjNov As ClsNovedad In lcolNovedades
                lenuIdTipoNov = lobjNov.ObjIdTipoNovedadByt.ObjValorPro
                ldecVlrNov = lobjNov.ObjValor_NovDec.ObjValorPro
                Select Case lenuIdTipoNov
                    Case EnuTipoNov.EnuCrAnApCap, EnuTipoNov.EnuCrAnApInt,
                            EnuTipoNov.EnuCrDctoCap, EnuTipoNov.EnuCrDctoInt,
                            EnuTipoNov.EnuCrPagoCap, EnuTipoNov.EnuCrPagoInt,
                            EnuTipoNov.EnuCrRetCre, EnuTipoNov.EnuCrRetFte,
                            EnuTipoNov.EnuCrRetIca, EnuTipoNov.EnuCrRetIva,
                            EnuTipoNov.EnuCrIvaGas
                        ldecCrNovs += ldecVlrNov
                    Case EnuTipoNov.EnuRCrPagoCap, EnuTipoNov.EnuRCrPagoInt,
                            EnuTipoNov.EnuRCrAnApCap, EnuTipoNov.EnuRCrAnApInt,
                            EnuTipoNov.EnuRCrDctoCap, EnuTipoNov.EnuRCrDctoInt,
                            EnuTipoNov.EnuRCrRetFte, EnuTipoNov.EnuRCrRetIva,
                            EnuTipoNov.EnuRCrRetIca, EnuTipoNov.EnuRCrRetCre,
                            EnuTipoNov.EnuRCrIvaGas
                        ldecDbNovs += ldecVlrNov
                    Case EnuTipoNov.EnuDbCap, EnuTipoNov.EnuDbInt, EnuTipoNov.EnuDbIva,
                            EnuTipoNov.EnuDbIvaInt
                        ldecDbNovs += ldecVlrNov
                    Case EnuTipoNov.EnuRDbCap, EnuTipoNov.EnuRDbIva, EnuTipoNov.EnuRDbInt,
                            EnuTipoNov.EnuRDbIvaInt
                        ldecCrNovs += ldecVlrNov
                End Select
                If lobjNov.ObjIdPredioAgrupador_NovStr.ObjValorPro <>
                        lobjFac.ObjIdPredioAgrupador_FacStr.ObjValorPro Then
                    lstrLinea = "En la Novedad número " & lobjNov.ObjIdNovedadShr.ToString &
                        " de la Factura " & lobjFac.StrNumeroFactura &
                        " el Predio Agupador no corresponde al de la Factura!"
                    SAdicioneErrorIntegridad(lstrLinea)
                End If
            Next
            If ldecDbNovs <> ldecDbItem Then
                lstrLinea = "En factura " & aobjItemfact.ObjIdFactura_ItemFactEnt.ObjValorPro &
                        " débitos en novedades difiere del débito en ItemFac " &
                        aobjItemfact.ObjIdItemFacturaShr.ObjValorPro
                SAdicioneErrorIntegridad(lstrLinea)
            End If
            If ldecCrNovs <> ldecCrItem Then
                lstrLinea = "En factura " & aobjItemfact.ObjIdFactura_ItemFactEnt.ObjValorPro &
                        " créditos en novedades difiere del crédito en ItemFac " &
                        aobjItemfact.ObjIdItemFacturaShr.ObjValorPro
                SAdicioneErrorIntegridad(lstrLinea)
            End If
        Else
            lstrLinea = "En factura " & lobjFac.ObjIdFacturaEnt.ObjValorPro &
                    " el ítem " & aobjItemfact.ObjIdItemFacturaShr.ObjValorPro &
                    " no generó novedades!"
            SAdicioneErrorIntegridad(lstrLinea)
        End If
    End Sub
    Private Sub SCompruebeRecs(astrFechaIni As String, astrFechaFin As String)
        Dim lstrCamposSelect = {ClsPrefijo_RecStr.SstrNombreCampoBd, ClsIdRecCajaEnt.SstrNombreCampoBd}
        Dim lstrIndice(,) = {{ClsPrefijo_RecStr.SstrNombreCampoBd, "ASC"},
                {ClsIdRecCajaEnt.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " & ClsFechaRecDtm.SstrNombreCampoBd &
                " >= " & astrFechaIni & " AND " & ClsFechaRecDtm.SstrNombreCampoBd &
                " <= " & astrFechaFin
        Dim ldtbIdRecibos = ClsPanorama.FdtbDataTable(ClsReciboCaja.SstrNombreTabla, lstrCamposSelect,
                    lstrIndice, lstrFiltro)
        Dim lobjRecibo As New ClsReciboCaja()
        ObjArgumentoEventoPan.EnuProceso = EnuProcesoDef.enuIntegrRec
        ObjArgumentoEventoPan.DblCantAProcesar = ldtbIdRecibos.Rows.Count
        ObjArgumentoEventoPan.DblCantProcesada = 0
        Dim i = 0
        RaiseEvent EvnInicio(Me, ObjArgumentoEventoPan)
        For Each ldrwIdRec As DataRow In ldtbIdRecibos.Rows
            i += 1
            Dim lstrPrefijo As String = ClsPanorama.FobjValorCampo(ldrwIdRec(ClsPrefijo_RecStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            Dim lentIdRec As Integer = ClsPanorama.FobjValorCampo(ldrwIdRec(ClsIdRecCajaEnt.SstrNombreCampoBd),
                    EnuTipoValor.enuInteger)
            SCompruebeRecibo(lobjRecibo, lstrPrefijo, lentIdRec)
            ObjArgumentoEventoPan.DblCantProcesada = i
            RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
            If ObjArgumentoEventoPan.BlnCancele Then
                Exit For
            End If
        Next
    End Sub
    Private Sub SCompruebeRecibo(aobjRecibo As ClsReciboCaja, astrPrefijoRec As String,
    aentIdRecibo As Integer)
        Dim lobjValorLlave() As Object = {GshrIdCarpeta, GshrIdCentroUtil, astrPrefijoRec, aentIdRecibo}
        Try
            With aobjRecibo
                Dim lstrLinea = String.Empty
                .SAbra(lobjValorLlave)
                If Not .BlnExiste Then
                    Throw New ErrorInesperadoPanLException("No existe Recibo de Caja con Id de Tabla." &
                            "  No se puede comprobar Integridad")
                End If
                Dim ldecValorRecibo As Decimal = .ObjValor_RecDec.ObjValorPro
                If ldecValorRecibo < 0 Then
                    lstrLinea = "Recibo de Caja " & .ObjIdRecCajaEnt.ToString & " con valor negativo"
                    SAdicioneErrorIntegridad(lstrLinea)
                Else
                    If aobjRecibo.ObjAnuladoBln.ObjValorPro AndAlso
                            Not IsNothing(aobjRecibo.ObjNotaReversionCr) Then
                        SCompruebeRecAnulado(aobjRecibo)
                    Else
                        SCompruebeNovRec(aobjRecibo)
                    End If
                End If
                SCompruebeMediosPago(aobjRecibo)
            End With
        Catch ex As PanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Sub
    Private Sub SCompruebeNovRec(aobjRecibo As ClsReciboCaja)
        Dim lstrLinea = String.Empty
        Dim ldecVlrAnticipo = 0D
        Dim lstrPrefijoRec As String = aobjRecibo.ObjPrefijo_RecStr.ObjValorPro
        Dim lentIdReci As Integer = aobjRecibo.ObjIdRecCajaEnt.ObjValorPro
        With aobjRecibo
            Dim ldtbNovedadesRec As DataTable = FdtbNovedadesRecibo(lstrPrefijoRec, lentIdReci)
            Dim lcolItemsRec = .ColItemsRecCaja
            Dim ldecSumaValorItems = 0D
            For Each lobjItemRec As ClsItemRecCaja In lcolItemsRec
                Dim ldecValorItem As Decimal = lobjItemRec.ObjValor_ItemRecDec.ObjValorPro
                Dim lenuTipoItemRec As EnuTipoItemRecCajaDef = lobjItemRec.ObjIdTipoItemRecByt.ObjValorPro
                Select Case lenuTipoItemRec
                    Case EnuTipoItemRecCajaDef.EnuAbonoCapital, EnuTipoItemRecCajaDef.EnuAbonoIntMora
                        ldecSumaValorItems += ldecValorItem
                    Case EnuTipoItemRecCajaDef.EnuAnticipo
                        ldecSumaValorItems += ldecValorItem
                        ldecVlrAnticipo += ldecValorItem
                End Select
                If lenuTipoItemRec <> EnuTipoItemRecCajaDef.EnuAnticipo Then
                    SVerifiqueIntegItemsRec(ldtbNovedadesRec, lobjItemRec)
                End If
            Next
            If .ObjValorAnticipoDec.ObjValorPro <> ldecVlrAnticipo Then
                lstrLinea = "En Recibo " & .ObjIdRecCajaEnt.ToString &
                        " difiere valor anticipo del valor del ítem"
                SAdicioneErrorIntegridad(lstrLinea)
            End If
            If String.IsNullOrEmpty(lstrLinea) AndAlso .ObjValor_RecDec.ObjValorPro <> ldecSumaValorItems Then
                lstrLinea = "En Recibo " & .ObjIdRecCajaEnt.ToString &
" difiere valor recibo del valor de los items"
                SAdicioneErrorIntegridad(lstrLinea)
            End If
        End With
    End Sub
    Private Sub SVerifiqueIntegItemsRec(adtbNovedades As DataTable, aobjItemRec As ClsItemRecCaja)
        Dim lstrMens = String.Empty
        Dim lstrFiltro = ClsIdItemDocOrigen_NovShr.SstrNombreCampoBd & " = " &
aobjItemRec.ObjIdItemRecCajaShr.ObjValorPro
        Dim ldrwNovedades As DataRow() = adtbNovedades.Select(lstrFiltro)
        If ldrwNovedades.Length > 0 Then
            Dim ldecVlrNovedades = 0D
            Dim lenuTipoItemRec As EnuTipoItemRecCajaDef = aobjItemRec.ObjIdTipoItemRecByt.ObjValorPro
            Dim lenuTipoNov As EnuTipoNov
            Dim lblnHayError As Boolean
            For Each ldrwNov As DataRow In ldrwNovedades
                lenuTipoNov = ClsPanorama.FobjValorCampo(ldrwNov(ClsIdTipoNovedadByt.SstrNombreCampoBd),
EnuTipoValor.enuInteger)
                Select Case lenuTipoNov
                    Case EnuTipoNov.EnuCrDctoCap
                        lblnHayError = (lenuTipoItemRec <> EnuTipoItemRecCajaDef.EnuDsctoCapital) AndAlso
    (lenuTipoItemRec <> EnuTipoItemRecCajaDef.EnuDsctoPP)
                        ldecVlrNovedades += ClsPanorama.FobjValorCampo(ldrwNov(
    ClsValor_NovDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
                    Case EnuTipoNov.EnuRCrDctoCap
                        lblnHayError = (lenuTipoItemRec <> EnuTipoItemRecCajaDef.EnuDsctoCapital) AndAlso
    (lenuTipoItemRec <> EnuTipoItemRecCajaDef.EnuDsctoPP)
                        ldecVlrNovedades -= ClsPanorama.FobjValorCampo(ldrwNov(
    ClsValor_NovDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
                    Case EnuTipoNov.EnuCrDctoInt
                        lblnHayError = (lenuTipoItemRec <> EnuTipoItemRecCajaDef.EnuDsctoIntMora)
                        ldecVlrNovedades += ClsPanorama.FobjValorCampo(ldrwNov(
    ClsValor_NovDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
                    Case EnuTipoNov.EnuRCrDctoInt
                        lblnHayError = (lenuTipoItemRec <> EnuTipoItemRecCajaDef.EnuDsctoIntMora)
                        ldecVlrNovedades -= ClsPanorama.FobjValorCampo(ldrwNov(
    ClsValor_NovDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
                    Case EnuTipoNov.EnuCrPagoCap
                        lblnHayError = (lenuTipoItemRec <> EnuTipoItemRecCajaDef.EnuAbonoCapital)
                        ldecVlrNovedades += ClsPanorama.FobjValorCampo(ldrwNov(
    ClsValor_NovDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
                    Case EnuTipoNov.EnuRCrPagoCap
                        lblnHayError = (lenuTipoItemRec <> EnuTipoItemRecCajaDef.EnuAbonoCapital)
                        ldecVlrNovedades -= ClsPanorama.FobjValorCampo(ldrwNov(
    ClsValor_NovDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
                    Case EnuTipoNov.EnuCrPagoInt
                        lblnHayError = (lenuTipoItemRec <> EnuTipoItemRecCajaDef.EnuAbonoIntMora)
                        ldecVlrNovedades += ClsPanorama.FobjValorCampo(ldrwNov(
    ClsValor_NovDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
                    Case EnuTipoNov.EnuRCrPagoInt
                        lblnHayError = (lenuTipoItemRec <> EnuTipoItemRecCajaDef.EnuAbonoIntMora)
                        ldecVlrNovedades -= ClsPanorama.FobjValorCampo(ldrwNov(
    ClsValor_NovDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
                    Case EnuTipoNov.EnuCrRetFte
                        lblnHayError = (lenuTipoItemRec <> EnuTipoItemRecCajaDef.EnuReteFuente)
                        ldecVlrNovedades += ClsPanorama.FobjValorCampo(ldrwNov(
    ClsValor_NovDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
                    Case EnuTipoNov.EnuRCrRetFte
                        lblnHayError = (lenuTipoItemRec <> EnuTipoItemRecCajaDef.EnuReteFuente)
                        ldecVlrNovedades -= ClsPanorama.FobjValorCampo(ldrwNov(
    ClsValor_NovDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
                    Case EnuTipoNov.EnuCrRetIca
                        lblnHayError = (lenuTipoItemRec <> EnuTipoItemRecCajaDef.EnuReteIca)
                        ldecVlrNovedades += ClsPanorama.FobjValorCampo(ldrwNov(
    ClsValor_NovDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
                    Case EnuTipoNov.EnuRCrRetIca
                        ldecVlrNovedades -= ClsPanorama.FobjValorCampo(ldrwNov(
    ClsValor_NovDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
                        lblnHayError = (lenuTipoItemRec <> EnuTipoItemRecCajaDef.EnuReteIca)
                    Case EnuTipoNov.EnuCrRetIva
                        lblnHayError = (lenuTipoItemRec <> EnuTipoItemRecCajaDef.EnuReteIva)
                        ldecVlrNovedades += ClsPanorama.FobjValorCampo(ldrwNov(
    ClsValor_NovDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
                    Case EnuTipoNov.EnuRCrRetIva
                        lblnHayError = (lenuTipoItemRec <> EnuTipoItemRecCajaDef.EnuReteIva)
                        ldecVlrNovedades -= ClsPanorama.FobjValorCampo(ldrwNov(
    ClsValor_NovDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
                    Case Else
                        lblnHayError = True
                End Select
                If lblnHayError Then
                    lstrMens = "En Recibo " & aobjItemRec.ObjIdRecCaja_ItemRecEnt.ObjValorPro &
" el tipo de movimiento del ítem " &
aobjItemRec.ObjIdItemRecCajaShr.ObjValorPro &
" es diferente al tipo de la Novedad generada."
                End If
            Next
            If ldecVlrNovedades <> aobjItemRec.ObjValor_ItemRecDec.ObjValorPro Then
                lstrMens = "En Recibo " & aobjItemRec.ObjIdRecCaja_ItemRecEnt.ObjValorPro & " el valor del ítem " &
aobjItemRec.ObjIdItemRecCajaShr.ObjValorPro & " es diferente al valor de la Novedad generada."
            End If
        Else
            Dim lobjRecCaja As ClsReciboCaja = aobjItemRec.ObjPadre
            If Not (aobjItemRec.ObjValor_ItemRecDec.ObjValorPro = 0 AndAlso lobjRecCaja.ObjAnuladoBln.ObjValorPro) Then
                lstrMens = "En Recibo " & aobjItemRec.ObjIdRecCaja_ItemRecEnt.ObjValorPro & " el ítem " &
aobjItemRec.ObjIdItemRecCajaShr.ObjValorPro & " no generó Novedad."
            End If
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            SAdicioneErrorIntegridad(lstrMens)
        End If
    End Sub
    Private Sub SCompruebeRecAnulado(aobjRec As ClsReciboCaja)
        Dim lstrMens = String.Empty
        Dim ldecVlrNovRec = 0D, ldecVlrNovNrrc = 0D
        For Each lobjNov As ClsNovedad In aobjRec.ColNovedades
            ldecVlrNovRec += lobjNov.ObjValor_NovDec.ObjValorPro
        Next
        For Each lobjNov As ClsNovedad In aobjRec.ObjNotaReversionCr.ColNovedades
            ldecVlrNovNrrc += lobjNov.ObjValor_NovDec.ObjValorPro
        Next
        For Each lobjNov As ClsNovedadAnticipo In aobjRec.ObjNotaReversionCr.ColNovedadesAnt
            If lobjNov.ObjIdTipoNovedad_NovAntByt.ObjValorPro = EnuTipoNov.EnuRDbAntApl Then
                ldecVlrNovNrrc -= lobjNov.ObjValor_NovAntDec.ObjValorPro
            End If
        Next
        If ldecVlrNovRec <> ldecVlrNovNrrc Then
            lstrMens = "En Recibo de Caja " & aobjRec.StrNumeroRecCaja & " anulado, elvalor de " &
                    "las Novedades difiere del valor de la novedades en La Nota de Reversión " &
                    "de Recibo de caja"
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            SAdicioneErrorIntegridad(lstrMens)
        End If
    End Sub
    Private Sub SCompruebeMediosPago(aobjRecibo As ClsReciboCaja)
        Dim lcolMediosPago As Collection = aobjRecibo.ColMediosPago
        Dim ldecVlrMedPag = 0D
        For Each lobjMedPag As ClsMedioPago In lcolMediosPago
            ldecVlrMedPag += lobjMedPag.ObjValor_MedPagoDec.ObjValorPro
        Next
        If ldecVlrMedPag <> aobjRecibo.ObjValor_RecDec.ObjValorPro Then
            Dim lstrMens = "En Recibo " & aobjRecibo.ObjIdRecCajaEnt.ObjValorPro &
                    "El valor de los Medios de Pago difiere del valor del Recibo de Caja"
            SAdicioneErrorIntegridad(lstrMens)
        End If
    End Sub
    Private Sub SCompruebeNotasCr(astrFechaIni As String, astrFechaFin As String)
        Dim lstrPrefijo As String, lentIdNcr As Integer
        Dim lstrCamposSelect = {ClsPrefijo_NotaCrStr.SstrNombreCampoBd, ClsIdNotaCrEnt.SstrNombreCampoBd}
        Dim lstrIndice(,) = {{ClsPrefijo_NotaCrStr.SstrNombreCampoBd, "ASC"},
                {ClsIdNotaCrEnt.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " & ClsFecha_NotaCrDtm.SstrNombreCampoBd &
                " >= " & astrFechaIni & " AND " & ClsFecha_NotaCrDtm.SstrNombreCampoBd &
                " <= " & astrFechaFin
        Dim ldtbIdDoc = ClsPanorama.FdtbDataTable(ClsNotaCr.SstrNombreTabla, lstrCamposSelect,
                    lstrIndice, lstrFiltro)
        Dim lobjDoc As New ClsNotaCr()
        ObjArgumentoEventoPan.EnuProceso = EnuProcesoDef.enuIntegrNcr
        ObjArgumentoEventoPan.DblCantAProcesar = ldtbIdDoc.Rows.Count
        ObjArgumentoEventoPan.DblCantProcesada = 0
        Dim i = 0
        RaiseEvent EvnInicio(Me, ObjArgumentoEventoPan)
        For Each ldrwIdDoc As DataRow In ldtbIdDoc.Rows
            i += 1
            lstrPrefijo = ClsPanorama.FobjValorCampo(ldrwIdDoc(ClsPrefijo_NotaCrStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            lentIdNcr = ClsPanorama.FobjValorCampo(ldrwIdDoc(ClsIdNotaCrEnt.SstrNombreCampoBd),
                    EnuTipoValor.enuInteger)
            SCompruebeNotaCr(lobjDoc, lstrPrefijo, lentIdNcr)
            ObjArgumentoEventoPan.DblCantProcesada = i
            RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
            If ObjArgumentoEventoPan.BlnCancele Then
                Exit For
            End If
        Next
    End Sub
    Private Sub SCompruebeNotaCr(aobjDoc As ClsNotaCr, astrPrefijoNcr As String,
    aentIdNCR As Integer)
        Dim lobjValorLlave() As Object = {GshrIdCarpeta, GshrIdCentroUtil, astrPrefijoNcr, aentIdNCR}
        Try
            With aobjDoc
                Dim lstrLinea = String.Empty
                .SAbra(lobjValorLlave)
                If Not .BlnExiste Then
                    Throw New ErrorInesperadoPanLException("No existe Nota Cr con Id de Tabla." &
                            "  No se puede comprobar Integridad")
                End If
                Dim ldecValorDoc As Decimal = .ObjValor_NotaCrDec.ObjValorPro
                If ldecValorDoc < 0 Then
                    lstrLinea = "Nota Cr " & .ObjIdNotaCrEnt.ToString & " con valor negativo"
                    SAdicioneErrorIntegridad(lstrLinea)
                Else
                    SCompruebeNovNCR(aobjDoc)
                End If
            End With
        Catch ex As PanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Sub
    Private Sub SCompruebeNovNCR(aobjDoc As ClsNotaCr)
        Dim lstrLinea = String.Empty
        With aobjDoc
            Dim ldecSumaValorItems = 0D
            For Each lobjItemNCR As ClsItemNotaCr In .ColItemsNotaCr
                Dim ldecValorItem As Decimal = lobjItemNCR.ObjValor_ItemNotaCrDec.ObjValorPro
                If Not (aobjDoc.ObjAnuladoBln.ObjValorPro) Then
                    ldecSumaValorItems += ldecValorItem
                    SVerifiqueIntegItemsNCR(lobjItemNCR)
                End If
            Next
            If String.IsNullOrEmpty(lstrLinea) AndAlso .ObjValor_NotaCrDec.ObjValorPro <> ldecSumaValorItems Then
                lstrLinea = "En Nota Crédito " & .ObjIdNotaCrEnt.ToString &
                        " difiere valor de la nota del valor de los items"
                SAdicioneErrorIntegridad(lstrLinea)
            End If
        End With
    End Sub
    Private Sub SVerifiqueIntegItemsNCR(aobjItemNCR As ClsItemNotaCr)
        Dim lstrMens = String.Empty
        Dim lobjNotaCr As ClsNotaCr = aobjItemNCR.ObjPadre
        If lobjNotaCr.ColNovedades.Count > 0 Then
            Dim ldecVlrNovedades = 0D
            Dim ldecVlrNov As Decimal
            Dim lenuTipoItemNCR As EnuTipoDescuentoDef = aobjItemNCR.ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro
            Dim lenuTipoNov As EnuTipoNov
            Dim lblnHayError As Boolean
            For Each lobjNov As ClsNovedad In lobjNotaCr.ColNovedades
                If lobjNov.ObjIdItemDocOrigen_NovShr.ObjValorPro = aobjItemNCR.ObjIdItemNotaCrShr.ObjValorPro Then
                    lenuTipoNov = lobjNov.ObjIdTipoNovedadByt.ObjValorPro
                    ldecVlrNov = lobjNov.ObjValor_NovDec.ObjValorPro
                    Select Case lenuTipoNov
                        Case EnuTipoNov.EnuCrDctoCap, EnuTipoNov.EnuRDbIva,
                                EnuTipoNov.EnuRDbCap, EnuTipoNov.EnuRDbInt,
                                EnuTipoNov.EnuCrIvaGas, EnuTipoNov.EnuRDbIvaInt
                            lblnHayError = (lenuTipoItemNCR <> EnuTipoDescuentoDef.EnuDsctoCapital) AndAlso
                                    (lenuTipoItemNCR <> EnuTipoDescuentoDef.EnuDsctoPP) AndAlso
                                    (lenuTipoItemNCR <> EnuTipoDescuentoDef.EnuDsctoIntMora) AndAlso
                                    (lenuTipoItemNCR <> EnuTipoDescuentoDef.EnuCancelaIva)
                            ldecVlrNovedades += ldecVlrNov
                        Case EnuTipoNov.EnuRCrDctoCap, EnuTipoNov.EnuRCrIvaGas
                            lblnHayError = (lenuTipoItemNCR <> EnuTipoDescuentoDef.EnuDsctoCapital) AndAlso
                                    (lenuTipoItemNCR <> EnuTipoDescuentoDef.EnuDsctoPP) AndAlso
                                    (lenuTipoItemNCR <> EnuTipoDescuentoDef.EnuCancelaIva)
                            ldecVlrNovedades -= ldecVlrNov
                        Case EnuTipoNov.EnuCrDctoInt
                            lblnHayError = (lenuTipoItemNCR <> EnuTipoDescuentoDef.EnuDsctoIntMora)
                            ldecVlrNovedades += ldecVlrNov
                        Case EnuTipoNov.EnuRCrDctoInt
                            lblnHayError = (lenuTipoItemNCR <> EnuTipoDescuentoDef.EnuDsctoIntMora)
                            ldecVlrNovedades -= ldecVlrNov
                        Case EnuTipoNov.EnuCrRetFte
                            lblnHayError = (lenuTipoItemNCR <> EnuTipoDescuentoDef.EnuReteFuente)
                            ldecVlrNovedades += ldecVlrNov
                        Case EnuTipoNov.EnuRCrRetFte
                            lblnHayError = (lenuTipoItemNCR <> EnuTipoDescuentoDef.EnuReteFuente)
                            ldecVlrNovedades -= ldecVlrNov
                        Case EnuTipoNov.EnuCrRetIca
                            lblnHayError = (lenuTipoItemNCR <> EnuTipoDescuentoDef.EnuReteIca)
                            ldecVlrNovedades += ldecVlrNov
                        Case EnuTipoNov.EnuRCrRetIca
                            lblnHayError = (lenuTipoItemNCR <> EnuTipoDescuentoDef.EnuReteIca)
                            ldecVlrNovedades -= ldecVlrNov
                        Case EnuTipoNov.EnuCrRetIva
                            lblnHayError = (lenuTipoItemNCR <> EnuTipoDescuentoDef.EnuReteIva)
                            ldecVlrNovedades += ldecVlrNov
                        Case EnuTipoNov.EnuRCrRetIva
                            lblnHayError = (lenuTipoItemNCR <> EnuTipoDescuentoDef.EnuReteIva)
                            ldecVlrNovedades -= ldecVlrNov
                        Case Else
                            lblnHayError = True
                    End Select
                    If lblnHayError Then
                        lstrMens = "La Nota Crédito " & aobjItemNCR.ObjIdNotaCr_ItemNotaCrEnt.ObjValorPro &
                                " el tipo de movimiento del ítem " &
                                aobjItemNCR.ObjIdItemNotaCrShr.ObjValorPro &
                                " es diferente al tipo de la Novedad generada." & lenuTipoNov.ToString & " " &
                                CType(lenuTipoNov, Byte).ToString
                    End If
                End If
            Next
            If ldecVlrNovedades <> aobjItemNCR.ObjValor_ItemNotaCrDec.ObjValorPro Then
                lstrMens = "La Nota Cr " & aobjItemNCR.ObjIdNotaCr_ItemNotaCrEnt.ObjValorPro &
                        " el valor del ítem " & aobjItemNCR.ObjIdItemNotaCrShr.ObjValorPro &
                        " es diferente al valor de la Novedad generada."
            End If
        Else
            lstrMens = "La Nota Crédito " & aobjItemNCR.ObjIdNotaCr_ItemNotaCrEnt.ObjValorPro &
                    " el ítem " & aobjItemNCR.ObjIdItemNotaCrShr.ObjValorPro &
                    " no generó Novedad."
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            SAdicioneErrorIntegridad(lstrMens)
        End If
    End Sub
    Private Sub SCompruebeNotasCon(astrFechaIni As String, astrFechaFin As String)
        Dim lstrPrefijo As String, lentIdNcon As Integer
        Dim lstrCamposSelect = {ClsPrefijo_NotaConStr.SstrNombreCampoBd, ClsIdNotaConEnt.SstrNombreCampoBd}
        Dim lstrIndice(,) = {{ClsPrefijo_NotaConStr.SstrNombreCampoBd, "ASC"},
                {ClsIdNotaConEnt.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " & ClsFecha_NotaConDtm.SstrNombreCampoBd &
                " >= " & astrFechaIni & " AND " & ClsFecha_NotaConDtm.SstrNombreCampoBd &
                " <= " & astrFechaFin
        Dim ldtbIdDoc = ClsPanorama.FdtbDataTable(ClsNotaCon.SstrNombreTabla, lstrCamposSelect,
                    lstrIndice, lstrFiltro)
        Dim lobjDoc As New ClsNotaCon()
        ObjArgumentoEventoPan.EnuProceso = EnuProcesoDef.enuIntegrNco
        ObjArgumentoEventoPan.DblCantAProcesar = ldtbIdDoc.Rows.Count
        ObjArgumentoEventoPan.DblCantProcesada = 0
        Dim i = 0
        RaiseEvent EvnInicio(Me, ObjArgumentoEventoPan)
        For Each ldrwIdDoc As DataRow In ldtbIdDoc.Rows
            i += 1
            lstrPrefijo = ClsPanorama.FobjValorCampo(ldrwIdDoc(ClsPrefijo_NotaConStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            lentIdNcon = ClsPanorama.FobjValorCampo(ldrwIdDoc(ClsIdNotaConEnt.SstrNombreCampoBd),
                    EnuTipoValor.enuInteger)
            SCompruebeNotaCon(lobjDoc, lstrPrefijo, lentIdNcon)
            ObjArgumentoEventoPan.DblCantProcesada = i
            RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
            If ObjArgumentoEventoPan.BlnCancele Then
                Exit For
            End If
        Next
    End Sub
    Private Sub SCompruebeNotaCon(aobjDoc As ClsNotaCon, astrPrefijoNCO As String,
    aentIdNCO As Integer)
        Dim lobjValorLlave() As Object = {GshrIdCarpeta, GshrIdCentroUtil, astrPrefijoNCO, aentIdNCO}
        Try
            With aobjDoc
                Dim lstrLinea = String.Empty
                .SAbra(lobjValorLlave)
                If Not .BlnExiste Then
                    Throw New ErrorInesperadoPanLException("No existe Nota Con con Id de Tabla." &
                            "  No se puede comprobar Integridad")
                End If
                Dim ldecValorDoc As Decimal = .ObjValor_NotaConDec.ObjValorPro
                If ldecValorDoc < 0 Then
                    lstrLinea = "Nota Con " & .ObjIdNotaConEnt.ToString & " con valor negativo"
                    SAdicioneErrorIntegridad(lstrLinea)
                Else
                    SCompruebeNovNCON(aobjDoc)
                End If
            End With
        Catch ex As PanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Sub
    Private Sub SCompruebeNovNCON(aobjDoc As ClsNotaCon)
        Dim lstrLinea = String.Empty
        Dim ldecVlrAnticipo = 0D
        Dim lstrPrefijoDoc As String = aobjDoc.ObjPrefijo_NotaConStr.ObjValorPro
        Dim lentIdDoc As Integer = aobjDoc.ObjIdNotaConEnt.ObjValorPro
        With aobjDoc
            Dim ldtbNovedadesNCON As DataTable = FdtbNovedadesNCON(lstrPrefijoDoc, lentIdDoc)
            Dim lcolItemsNCON = .ColItemsNotaCon
            Dim ldecSumaValorItems = 0D
            For Each lobjItemNCON As ClsItemNotaCon In lcolItemsNCON
                Dim ldecValorItem As Decimal = lobjItemNCON.ObjValor_ItemNotaConDec.ObjValorPro
                ldecSumaValorItems += ldecValorItem
                SVerifiqueIntegItemsNCON(ldtbNovedadesNCON, lobjItemNCON)
            Next
            If String.IsNullOrEmpty(lstrLinea) AndAlso .ObjValor_NotaConDec.ObjValorPro <> ldecSumaValorItems Then
                lstrLinea = "En Nota Contable " & .ObjIdNotaConEnt.ToString &
" difiere valor de la nota del valor de los items"
                SAdicioneErrorIntegridad(lstrLinea)
            End If
        End With
    End Sub
    Private Sub SVerifiqueIntegItemsNCON(adtbNovedades As DataTable, aobjItemNCON As ClsItemNotaCon)
        Dim lstrMens = String.Empty
        Dim lstrFiltro = ClsIdItemDocOrigen_NovShr.SstrNombreCampoBd & " = " &
                aobjItemNCON.ObjIdItemNotaConShr.ObjValorPro
        Dim ldrwNovedades As DataRow() = adtbNovedades.Select(lstrFiltro)
        If ldrwNovedades.Length > 0 Then
            Dim ldecVlrNov As Decimal, ldecTotVlrNov = 0D
            Dim lenuTipoItemNCON As EnuTipoItemNotaConDef = aobjItemNCON.ObjIdTipoItemNotaConByt.ObjValorPro
            Dim lenuTipoNov As EnuTipoNov
            Dim lblnHayError As Boolean
            For Each ldrwNov As DataRow In ldrwNovedades
                ldecVlrNov = ClsPanorama.FobjValorCampo(ldrwNov(
ClsValor_NovDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
                lenuTipoNov = ClsPanorama.FobjValorCampo(ldrwNov(ClsIdTipoNovedadByt.SstrNombreCampoBd),
EnuTipoValor.enuInteger)
                Select Case lenuTipoNov
                    Case EnuTipoNov.EnuCrDctoCap
                        lblnHayError = (lenuTipoItemNCON <> EnuTipoItemNotaConDef.EnuDsctoPP)
                        ldecTotVlrNov += ldecVlrNov
                    Case EnuTipoNov.EnuCrAnApCap
                        lblnHayError = (lenuTipoItemNCON <> EnuTipoItemNotaConDef.EnuAplicaAntCap)
                        ldecTotVlrNov += ldecVlrNov
                    Case EnuTipoNov.EnuCrAnApInt
                        lblnHayError = (lenuTipoItemNCON <> EnuTipoItemNotaConDef.EnuAplicaAntInt)
                        ldecTotVlrNov += ldecVlrNov
                    Case EnuTipoNov.EnuCrRetFte
                        lblnHayError = (lenuTipoItemNCON <> EnuTipoItemNotaConDef.EnuReteFuente)
                        ldecTotVlrNov += ldecVlrNov
                    Case EnuTipoNov.EnuCrRetIca
                        lblnHayError = (lenuTipoItemNCON <> EnuTipoItemNotaConDef.EnuReteIca)
                        ldecTotVlrNov += ldecVlrNov
                    Case EnuTipoNov.EnuCrRetIva
                        lblnHayError = (lenuTipoItemNCON <> EnuTipoItemNotaConDef.EnuReteIva)
                        ldecTotVlrNov += ldecVlrNov
                    Case EnuTipoNov.EnuRCrDctoCap
                        lblnHayError = (lenuTipoItemNCON <> EnuTipoItemNotaConDef.EnuDsctoPP)
                        ldecTotVlrNov -= ldecVlrNov
                    Case EnuTipoNov.EnuRCrAnApCap
                        lblnHayError = (lenuTipoItemNCON <> EnuTipoItemNotaConDef.EnuAplicaAntCap)
                        ldecTotVlrNov -= ldecVlrNov
                    Case EnuTipoNov.EnuRCrAnApInt
                        lblnHayError = (lenuTipoItemNCON <> EnuTipoItemNotaConDef.EnuAplicaAntInt)
                        ldecTotVlrNov -= ldecVlrNov
                    Case EnuTipoNov.EnuRCrRetFte
                        lblnHayError = (lenuTipoItemNCON <> EnuTipoItemNotaConDef.EnuReteFuente)
                        ldecTotVlrNov -= ldecVlrNov
                    Case EnuTipoNov.EnuRCrRetIca
                        lblnHayError = (lenuTipoItemNCON <> EnuTipoItemNotaConDef.EnuReteIca)
                        ldecTotVlrNov -= ldecVlrNov
                    Case EnuTipoNov.EnuRCrRetIva
                        lblnHayError = (lenuTipoItemNCON <> EnuTipoItemNotaConDef.EnuReteIva)
                        ldecTotVlrNov -= ldecVlrNov
                    Case Else
                        lblnHayError = True
                End Select
                If lblnHayError Then
                    lstrMens = "La Nota Contable " & aobjItemNCON.ObjIdNotaCon_ItemNotaConEnt.ObjValorPro &
" el tipo de movimiento del ítem " &
aobjItemNCON.ObjIdItemNotaConShr.ObjValorPro &
" es diferente al tipo de la Novedad generada." & lenuTipoNov.ToString & " " &
CType(lenuTipoNov, Byte).ToString
                    SAdicioneErrorIntegridad(lstrMens)
                End If
            Next
            If ldecTotVlrNov <> aobjItemNCON.ObjValor_ItemNotaConDec.ObjValorPro Then
                lstrMens = "La Nota Contable " & aobjItemNCON.ObjIdNotaCon_ItemNotaConEnt.ObjValorPro &
" el valor del ítem " & aobjItemNCON.ObjIdItemNotaConShr.ObjValorPro &
" es diferente al valor de la Novedad generada."
            End If
        Else
            lstrMens = "La Nota Contable " & aobjItemNCON.ObjIdNotaCon_ItemNotaConEnt.ObjValorPro &
" el ítem " & aobjItemNCON.ObjIdItemNotaConShr.ObjValorPro &
" no generó Novedad."
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            SAdicioneErrorIntegridad(lstrMens)
        End If
    End Sub
    Private Sub SCompruebeNotasDb(astrFechaIni As String, astrFechaFin As String)
        Dim lstrPrefijo As String, lentIdNdb As Integer
        Dim lstrCamposSelect = {ClsPrefijo_NotaDbStr.SstrNombreCampoBd,
                ClsIdNotaDbEnt.SstrNombreCampoBd}
        Dim lstrIndice(,) = {{ClsPrefijo_NotaDbStr.SstrNombreCampoBd, "ASC"},
                {ClsIdNotaDbEnt.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " &
                ClsFecha_NotaDbDtm.SstrNombreCampoBd & " >= " & astrFechaIni & " AND " &
                ClsFecha_NotaDbDtm.SstrNombreCampoBd & " <= " & astrFechaFin
        Dim ldtbIdDoc = ClsPanorama.FdtbDataTable(ClsNotaDb.SstrNombreTabla, lstrCamposSelect,
                    lstrIndice, lstrFiltro)
        Dim lobjDoc As New ClsNotaDb()
        ObjArgumentoEventoPan.EnuProceso = EnuProcesoDef.enuIntegrNdb
        ObjArgumentoEventoPan.DblCantAProcesar = ldtbIdDoc.Rows.Count
        ObjArgumentoEventoPan.DblCantProcesada = 0
        Dim i = 0
        RaiseEvent EvnInicio(Me, ObjArgumentoEventoPan)
        For Each ldrwIdDoc As DataRow In ldtbIdDoc.Rows
            i += 1
            lstrPrefijo = ClsPanorama.FobjValorCampo(ldrwIdDoc(ClsPrefijo_NotaDbStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            lentIdNdb = ClsPanorama.FobjValorCampo(ldrwIdDoc(ClsIdNotaDbEnt.SstrNombreCampoBd),
                    EnuTipoValor.enuInteger)
            SCompruebeNotaDb(lobjDoc, lstrPrefijo, lentIdNdb)
            ObjArgumentoEventoPan.DblCantProcesada = i
            RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
            If ObjArgumentoEventoPan.BlnCancele Then
                Exit For
            End If
        Next
    End Sub
    Private Sub SCompruebeNotaDb(aobjDoc As ClsNotaDb, astrPrefNdb As String,
            aentIdNdb As Integer)
        Dim lobjValorLlave() As Object = {GshrIdCarpeta, GshrIdCentroUtil, astrPrefNdb, aentIdNdb}
        Try
            With aobjDoc
                Dim lstrLinea = String.Empty
                .SAbra(lobjValorLlave)
                If Not .BlnExiste Then
                    Throw New ErrorInesperadoPanLException("No existe Nota Db con Id de Tabla." &
                            "  No se puede comprobar Integridad")
                End If
                Dim ldecValorDoc As Decimal = .ObjValor_NotaDbDec.ObjValorPro
                If ldecValorDoc < 0 Then
                    lstrLinea = "Nota Db " & .ObjIdNotaDbEnt.ToString & " con valor negativo"
                    SAdicioneErrorIntegridad(lstrLinea)
                Else
                    SCompruebeNovNdb(aobjDoc)
                End If
            End With
        Catch ex As PanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Sub
    Private Sub SCompruebeNovNdb(aobjDoc As ClsNotaDb)
        Dim lstrLinea = String.Empty
        Dim lstrPrefijoDoc As String = aobjDoc.ObjPrefijo_NotaDbStr.ObjValorPro
        Dim lentIdDoc As Integer = aobjDoc.ObjIdNotaDbEnt.ObjValorPro
        With aobjDoc
            Dim ldtbNovedadesNdb As DataTable = FdtbNovedadesNdb(lstrPrefijoDoc, lentIdDoc)
            Dim lcolItemsNdb = .ColItemsNotaDb
            Dim ldecSumaValorItems = 0D
            For Each lobjItemNdb As ClsItemNotaDb In lcolItemsNdb
                Dim ldecValorItem As Decimal = lobjItemNdb.ObjValor_ItemNotaDbDec.ObjValorPro
                ldecSumaValorItems += ldecValorItem
                SVerifiqueIntegItemsNdb(ldtbNovedadesNdb, lobjItemNdb)
            Next
            If String.IsNullOrEmpty(lstrLinea) AndAlso .ObjValor_NotaDbDec.ObjValorPro <>
                    ldecSumaValorItems Then
                lstrLinea = "En Nota Débito " & .ObjIdNotaDbEnt.ToString &
                        " difiere valor de la nota del valor de los items"
                SAdicioneErrorIntegridad(lstrLinea)
            End If
        End With
    End Sub
    Private Sub SVerifiqueIntegItemsNdb(adtbNovedades As DataTable,
        aobjItemNdb As ClsItemNotaDb)
        Dim lstrMens = String.Empty
        Dim lstrFiltro = ClsIdItemDocOrigen_NovShr.SstrNombreCampoBd & " = " &
                aobjItemNdb.ObjIdItemNotaDbShr.ObjValorPro
        Dim ldrwNovedades As DataRow() = adtbNovedades.Select(lstrFiltro)
        If ldrwNovedades.Length > 0 Then
            Dim ldecVlrNovedades = 0D
            Dim ldecVlrNov As Decimal
            Dim lenuTipoNov As EnuTipoNov
            For Each ldrwNov As DataRow In ldrwNovedades
                lenuTipoNov = ClsPanorama.FobjValorCampo(ldrwNov(
                        ClsIdTipoNovedadByt.SstrNombreCampoBd),
                        EnuTipoValor.enuInteger)
                ldecVlrNov = ClsPanorama.FobjValorCampo(ldrwNov(
                        ClsValor_NovDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
                If lenuTipoNov <> EnuTipoNov.EnuDbInt AndAlso (lenuTipoNov <>
                        EnuTipoNov.EnuDbIva AndAlso aobjItemNdb.ObjTarifaIva_ItemNotaDbDbl.
                        ObjValorPro = 0) Then
                    lstrMens = "La Nota Débito " &
                            aobjItemNdb.ObjIdNotaDb_ItemNotaDbEnt.ObjValorPro &
                            " el tipo de movimiento del ítem " &
                            aobjItemNdb.ObjIdItemNotaDbShr.ObjValorPro &
                            " es diferente al tipo de la Novedad generada." &
                            lenuTipoNov.ToString & " " & CType(lenuTipoNov, Byte).ToString
                End If
                ldecVlrNovedades += ldecVlrNov
            Next
            If ldecVlrNovedades <> aobjItemNdb.ObjValor_ItemNotaDbDec.ObjValorPro Then
                lstrMens = "La Nota Db " & aobjItemNdb.ObjIdNotaDb_ItemNotaDbEnt.ObjValorPro &
                        " el valor del ítem " & aobjItemNdb.ObjIdItemNotaDbShr.ObjValorPro &
                        " es diferente al valor de la Novedad generada."
            End If
        Else
            If Not aobjItemNdb.ObjAnuladoBln.ObjValorPro AndAlso
                    aobjItemNdb.ObjValor_ItemNotaDbDec.ObjValorPro <> 0 Then
                lstrMens = "La Nota Débito " & aobjItemNdb.ObjIdNotaDb_ItemNotaDbEnt.ObjValorPro &
                    " el ítem " & aobjItemNdb.ObjIdItemNotaDbShr.ObjValorPro &
                    " no generó Novedad."
            End If
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            SAdicioneErrorIntegridad(lstrMens)
        End If
    End Sub
    Private Sub SCompruebeAnticipos(astrFechaIni As String, astrFechaFin As String)
        Dim lentIdAnt As Integer
        Dim lstrCamposSelect = {ClsIdAnticipoEnt.SstrNombreCampoBd}
        Dim lstrIndice(,) = {{ClsIdAnticipoEnt.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " &
                ClsFechaAnticipoDtm.SstrNombreCampoBd & " >= " & astrFechaIni & " AND " &
                ClsFechaAnticipoDtm.SstrNombreCampoBd & " <= " & astrFechaFin
        Dim ldtbIdDoc = ClsPanorama.FdtbDataTable(ClsAnticipo.SstrNombreTabla, lstrCamposSelect,
                    lstrIndice, lstrFiltro)
        ObjArgumentoEventoPan.EnuProceso = EnuProcesoDef.enuInteAnt
        ObjArgumentoEventoPan.DblCantAProcesar = ldtbIdDoc.Rows.Count
        ObjArgumentoEventoPan.DblCantProcesada = 0
        Dim i = 0
        RaiseEvent EvnInicio(Me, ObjArgumentoEventoPan)
        Dim ldecDebitosNovs As Decimal, ldecCreditosNovs = 0D
        Dim lstrMens As String
        Dim lobjAnt As New ClsAnticipo(EnuModoInstanciaObjDef.enuUnico)
        For Each ldrwDoc As DataRow In ldtbIdDoc.Rows
            i += 1
            lentIdAnt = ClsPanorama.FobjValorCampo(ldrwDoc(ClsIdAnticipoEnt.SstrNombreCampoBd),
                    EnuTipoValor.enuInteger)
            lobjAnt.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lentIdAnt})
            lstrMens = String.Empty
            If lobjAnt.ObjDebitos_AntDec.ObjValorPro > lobjAnt.ObjCreditos_AntDec.ObjValorPro Then
                lstrMens = "En Anticipo " & lobjAnt.ObjIdAnticipoEnt.ToString &
                        " los Débitos son mayores a los Créditos!"
                SAdicioneErrorIntegridad(lstrMens)
            End If
            ldecDebitosNovs = FdecDbitosNov(lobjAnt.ColNovedadesAnt, ldecCreditosNovs)
            If ldecCreditosNovs <> lobjAnt.ObjCreditos_AntDec.ObjValorPro Then
                lstrMens = "En Anticipo " & lobjAnt.ObjIdAnticipoEnt.ToString &
                        " los Créditos tienen problema de integridad"
                SAdicioneErrorIntegridad(lstrMens)
            End If
            If ldecDebitosNovs <> lobjAnt.ObjDebitos_AntDec.ObjValorPro Then
                lstrMens = "En Anticipo " & lobjAnt.ObjIdAnticipoEnt.ToString &
                        " los Débitos tienen problema de integridad"
                SAdicioneErrorIntegridad(lstrMens)
            End If
            If ldecCreditosNovs <> lobjAnt.ObjCreditos_AntDec.ObjValorPro Then
                lstrMens = "En Anticipo " & lobjAnt.ObjIdAnticipoEnt.ToString &
                        " los Débitos tienen problema de integridad"
                SAdicioneErrorIntegridad(lstrMens)
            End If
            ObjArgumentoEventoPan.DblCantProcesada = i
            RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
            If ObjArgumentoEventoPan.BlnCancele Then
                Exit For
            End If
        Next
    End Sub
    Private Sub SCompruebeNotasRCr(astrFechaIni As String, astrFechaFin As String)
        Dim lentIdNrrc As Integer, lstrPrefNrrc As String
        Dim lstrCamposSelect = {ClsPrefijo_NotaReversaCrStr.SstrNombreCampoBd,
                ClsIdNotaReversaCrEnt.SstrNombreCampoBd}
        Dim lstrIndice(,) = {{ClsPrefijo_NotaReversaCrStr.SstrNombreCampoBd, "ASC"},
                {ClsIdNotaReversaCrEnt.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " &
                ClsFecha_NotaReversaCrDtm.SstrNombreCampoBd & " >= " & astrFechaIni & " AND " &
                ClsFecha_NotaReversaCrDtm.SstrNombreCampoBd & " <= " & astrFechaFin
        Dim ldtbIdDoc = ClsPanorama.FdtbDataTable(ClsNotaReversionCr.SstrNombreTabla, lstrCamposSelect,
                    lstrIndice, lstrFiltro)
        Dim lobjDoc As New ClsNotaReversionCr()
        ObjArgumentoEventoPan.EnuProceso = EnuProcesoDef.enuInteAnt
        ObjArgumentoEventoPan.DblCantAProcesar = ldtbIdDoc.Rows.Count
        ObjArgumentoEventoPan.DblCantProcesada = 0
        Dim i = 0
        RaiseEvent EvnInicio(Me, ObjArgumentoEventoPan)
        Dim lstrMens = String.Empty
        For Each ldrwDoc As DataRow In ldtbIdDoc.Rows
            i += 1
            Dim ldecVlrDoc = 0D, ldecVlrCrDoc = 0D, ldecValorRev = 0D, ldecVlrCrRev = 0D
            Dim ldecVlrAntApl = 0D, ldecVlrAntRev = 0D
            lstrPrefNrrc = ClsPanorama.FobjValorCampo(ldrwDoc(ClsPrefijo_NotaReversaCrStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            lentIdNrrc = ClsPanorama.FobjValorCampo(ldrwDoc(ClsIdNotaReversaCrEnt.SstrNombreCampoBd),
                    EnuTipoValor.enuInteger)
            lobjDoc.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefNrrc, lentIdNrrc})
            ldecVlrDoc = FdecVlrDoc(lobjDoc, ldecVlrCrDoc, ldecVlrAntApl)
            ldecValorRev = FdecValorRevNotaRCr(lobjDoc, ldecVlrCrRev, ldecVlrAntRev)
            If ldecVlrDoc <> ldecValorRev OrElse ldecVlrCrDoc <> ldecVlrCrRev OrElse
                    ldecVlrAntApl <> ldecVlrAntRev Then
                lstrMens = "El valor de la Nota Reversión de Cr " &
                        lobjDoc.ObjIdNotaReversaCrEnt.ToString &
                                " es diferente al valor del Documento Reversado!"
            End If
            If Not String.IsNullOrEmpty(lstrMens) Then
                SAdicioneErrorIntegridad(lstrMens)
                lstrMens = String.Empty
            End If
            ObjArgumentoEventoPan.DblCantProcesada = i
            RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
            If ObjArgumentoEventoPan.BlnCancele Then
                Exit For
            End If
        Next
    End Sub
    Private Shared Function FdecValorRevNotaRCr(aobjNotaRCr As ClsNotaReversionCr,
            ByRef adecCrRev As Decimal, ByRef adecAntRev As Decimal) As Decimal
        Dim ldecVlrRev = 0D, ldecVlrCrRev = 0D, ldecVlrAntRev = 0D
        For Each lobjNov As ClsNovedad In aobjNotaRCr.ColNovedades
            If lobjNov.ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuRCrPagoCap OrElse
                    lobjNov.ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuRCrPagoInt Then
                ldecVlrRev += lobjNov.ObjValor_NovDec.ObjValorPro
            ElseIf (lobjNov.ObjIdTipoNovedadByt.ObjValorPro >= EnuTipoNov.EnuRCrDctoCap AndAlso
                    lobjNov.ObjIdTipoNovedadByt.ObjValorPro <= EnuTipoNov.EnuRCrRetCre) OrElse
                    lobjNov.ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuRCrIvaGas Then
                ldecVlrCrRev += lobjNov.ObjValor_NovDec.ObjValorPro
            ElseIf lobjNov.ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuRCrAnApCap OrElse
                    lobjNov.ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuRCrAnApInt Then
                ldecVlrAntRev += lobjNov.ObjValor_NovDec.ObjValorPro
            End If
        Next
        For Each lobjNov As ClsNovedadAnticipo In aobjNotaRCr.ColNovedadesAnt
            If lobjNov.ObjIdTipoNovedad_NovAntByt.ObjValorPro = EnuTipoNov.EnuRCrAntRec Then
                ldecVlrRev += lobjNov.ObjValor_NovAntDec.ObjValorPro
            End If
        Next
        adecCrRev = ldecVlrCrRev
        adecAntRev = ldecVlrAntRev
        Return ldecVlrRev
    End Function
    Private Shared Function FdecDbitosNov(acolNovAnt As Collection,
            ByRef adecCreditosNov As Decimal) As Decimal
        Dim ldecNov As Decimal, ldecCr = 0D, ldecDb = 0D, lenuTipNov As EnuTipoNov
        For Each lobjNonAnt As ClsNovedadAnticipo In acolNovAnt
            lenuTipNov = lobjNonAnt.ObjIdTipoNovedad_NovAntByt.ObjValorPro
            ldecNov = lobjNonAnt.ObjValor_NovAntDec.ObjValorPro
            Select Case lenuTipNov
                Case EnuTipoNov.EnuCrAntRec, EnuTipoNov.EnuRDbAntDev,
                        EnuTipoNov.EnuRDbAntApl
                    ldecCr += ldecNov
                Case EnuTipoNov.EnuRCrAntRec, EnuTipoNov.EnuDbAntDev,
                        EnuTipoNov.EnuDbAntApl
                    ldecDb += ldecNov
            End Select
        Next
        adecCreditosNov = ldecCr
        Return ldecDb
    End Function
    Private Shared Function FdecVlrDoc(aobjNotaRCr As ClsNotaReversionCr,
            ByRef adecVlrCrDoc As Decimal, ByRef adecVlrAntApl As Decimal) As Decimal
        Dim ldecVlrDoc = 0D, ldecVlrCrDoc = 0D, lcolNovs As Collection
        If aobjNotaRCr.ObjTipoDocReversadoByt.ObjValorPro = EnuDocReversado.EnuReciboC Then
            Dim lobjRec As ClsReciboCaja = aobjNotaRCr.ObjDocReversado
            lcolNovs = lobjRec.ColNovedades
            If Not IsNothing(lobjRec.ObjAnticipo) Then
                For Each lobjNovAnt As ClsNovedadAnticipo In lobjRec.ObjAnticipo.ColNovedadesAnt
                    If lobjNovAnt.ObjIdTipoNovedad_NovAntByt.ObjValorPro = EnuTipoNov.EnuCrAntRec Then
                        ldecVlrDoc += lobjNovAnt.ObjValor_NovAntDec.ObjValorPro
                    End If
                Next
                adecVlrAntApl = FdecVrAntApl(lobjRec.ObjAnticipo, adecVlrCrDoc)
            End If
        Else
            Dim lobjNcr As ClsNotaCr = aobjNotaRCr.ObjDocReversado
            lcolNovs = lobjNcr.ColNovedades
        End If
        For Each lobjNov As ClsNovedad In lcolNovs
            If lobjNov.ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuCrPagoCap OrElse
                    lobjNov.ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuCrPagoInt Then
                ldecVlrDoc += lobjNov.ObjValor_NovDec.ObjValorPro
            ElseIf (lobjNov.ObjIdTipoNovedadByt.ObjValorPro >= EnuTipoNov.EnuCrDctoCap AndAlso
                    lobjNov.ObjIdTipoNovedadByt.ObjValorPro <= EnuTipoNov.EnuCrRetCre) OrElse
                    lobjNov.ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuCrIvaGas Then
                ldecVlrCrDoc += lobjNov.ObjValor_NovDec.ObjValorPro
            End If
        Next
        adecVlrCrDoc += ldecVlrCrDoc
        Return ldecVlrDoc
    End Function
    Private Shared Function FdecVrAntApl(aobjAnticipo As ClsAnticipo, ByRef adecVlrCr As Decimal) As Decimal
        Dim ldecAntApl = 0D, ldecVlrCrNota As Decimal
        For Each lobjNotCon As ClsNotaCon In aobjAnticipo.ColNotasCon
            ldecAntApl += FdecVlrAntAplPorNota(lobjNotCon, ldecVlrCrNota)
            adecVlrCr += ldecVlrCrNota
        Next
        Return ldecAntApl
    End Function
    Private Shared Function FdecVlrAntAplPorNota(aobjNotaCon As ClsNotaCon, ByRef adecVlrCr As Decimal) As Decimal
        Dim ldecAntApl = 0D
        For Each lobjNov As ClsNovedad In aobjNotaCon.ColNovedades
            If lobjNov.ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuCrAnApCap OrElse
                    lobjNov.ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuCrAnApInt Then
                ldecAntApl += lobjNov.ObjValor_NovDec.ObjValorPro
            ElseIf lobjNov.ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuCrDctoCap Then
                adecVlrCr += lobjNov.ObjValor_NovDec.ObjValorPro
            End If
        Next
        Return ldecAntApl
    End Function
    Private Shared Function FdtbNovedadesRecibo(astrPrefijoRec As String, aentIdRecibo As Integer) As DataTable
        Dim ldtbNovRec As DataTable
        Dim lstrIndice = {{ClsIdTipoDocOrigenByt.SstrNombreCampoBd, "ASC"},
                {ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd, "ASC"},
                {ClsIdDocOrigenEnt.SstrNombreCampoBd, "ASC"},
                {ClsPrefijoFact_NovStr.SstrNombreCampoBd, "ASC"},
                {ClsIdFactura_NovEnt.SstrNombreCampoBd, "ASC"},
                {ClsIdNovedadShr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " &
                ClsIdTipoDocOrigenByt.SstrNombreCampoBd & " = " & EnuTipoDocOri.EnuReciboCaja &
                " AND " & ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd & " = '" & astrPrefijoRec &
                "' AND " & ClsIdDocOrigenEnt.SstrNombreCampoBd & " = " & aentIdRecibo
        Dim lstrCamposSelect() = {"*"}
        ldtbNovRec = ClsPanorama.FdtbDataTable(ClsNovedad.SstrNombreTabla, lstrCamposSelect, lstrIndice,
                lstrFiltro)
        Return ldtbNovRec
    End Function
    Private Shared Function FdtbNovedadesNdb(astrPrefijoNdb As String,
                aentIdNdb As Integer) As DataTable
        Dim ldtbNovNdb As DataTable
        Dim lstrIndice = {{ClsIdTipoDocOrigenByt.SstrNombreCampoBd, "ASC"},
                {ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd, "ASC"},
                {ClsIdDocOrigenEnt.SstrNombreCampoBd, "ASC"},
                {ClsPrefijoFact_NovStr.SstrNombreCampoBd, "ASC"},
                {ClsIdFactura_NovEnt.SstrNombreCampoBd, "ASC"},
                {ClsIdNovedadShr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " &
                ClsIdTipoDocOrigenByt.SstrNombreCampoBd & " = " &
                EnuTipoDocOri.EnuNotaDb & " AND " &
                ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd & " = '" & astrPrefijoNdb &
                "' AND " & ClsIdDocOrigenEnt.SstrNombreCampoBd & " = " & aentIdNdb
        Dim lstrCamposSelect() = {"*"}
        ldtbNovNdb = ClsPanorama.FdtbDataTable(ClsNovedad.SstrNombreTabla, lstrCamposSelect, lstrIndice,
                lstrFiltro)
        Return ldtbNovNdb
    End Function
    Private Shared Function FdtbNovedadesNCON(astrPrefijoNCON As String, aentIdNCON As Integer) As DataTable
        Dim ldtbNovNCON As DataTable
        Dim lstrIndice = {{StrCampoCarpeta, "ASC"},
                {StrCampoCentroUtil, "ASC"},
                {ClsIdTipoDocOrigenByt.SstrNombreCampoBd, "ASC"},
                {ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd, "ASC"},
                {ClsIdDocOrigenEnt.SstrNombreCampoBd, "ASC"},
                {ClsPrefijoFact_NovStr.SstrNombreCampoBd, "ASC"},
                {ClsIdFactura_NovEnt.SstrNombreCampoBd, "ASC"},
                {ClsIdNovedadShr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " & ClsIdTipoDocOrigenByt.SstrNombreCampoBd &
                " = " & EnuTipoDocOri.EnuNotaCon & " AND " &
                ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd & " = '" & astrPrefijoNCON & "' AND " &
                ClsIdDocOrigenEnt.SstrNombreCampoBd & " = " & aentIdNCON
        Dim lstrCamposSelect() = {"*"}
        ldtbNovNCON = ClsPanorama.FdtbDataTable(ClsNovedad.SstrNombreTabla, lstrCamposSelect, lstrIndice,
lstrFiltro)
        Return ldtbNovNCON
    End Function
    Private Sub SCompruebeEstadoCtas(astrFecIni As String, astrFecFin As String)
        Dim lstrTabla = ClsEstadoCuenta.SstrNombreTabla
        Dim lstrCamposSel As String() = {ClsIdEstadoCuentaEnt.SstrNombreCampoBd}
        Dim lstrFiltro As String = StrFiltroUbicacion & " AND " &
                ClsFechaEstadoDtm.SstrNombreCampoBd & " >= " & astrFecIni & " AND " &
                ClsFechaEstadoDtm.SstrNombreCampoBd & " <= " & astrFecFin
        Dim lstrOrden As String(,) = {{ClsIdEstadoCuentaEnt.SstrNombreCampoBd, "ASC"}}
        Dim ldtbEstCta = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamposSel, lstrOrden, lstrFiltro)
        Dim lentIdEstCta As Integer, i = 0
        ObjArgumentoEventoPan.EnuProceso = EnuProcesoDef.enuInteEstadoCta
        ObjArgumentoEventoPan.DblCantAProcesar = ldtbEstCta.Rows.Count
        ObjArgumentoEventoPan.DblCantProcesada = 0
        RaiseEvent EvnInicio(Me, ObjArgumentoEventoPan)
        Dim lobjEstaCta As New ClsEstadoCuenta(EnuModoInstanciaObjDef.enuUnico, True)
        For Each ldrwCta As DataRow In ldtbEstCta.Rows
            lentIdEstCta = ClsPanorama.FobjValorCampo(ldrwCta(ClsIdEstadoCuentaEnt.SstrNombreCampoBd),
                    EnuTipoValor.enuInteger)
            lobjEstaCta.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lentIdEstCta})
            SCompruebeEstCta(lobjEstaCta)
            i += 1
            ObjArgumentoEventoPan.DblCantProcesada = i
            RaiseEvent EvnAvance(Me, ObjArgumentoEventoPan)
            If ObjArgumentoEventoPan.BlnCancele Then
                Exit For
            End If
        Next
    End Sub
    Private Sub SCompruebeEstCta(aobjEstaCta As ClsEstadoCuenta)
        Dim ldecDedudaCap = 0D, ldecDeudaInt = 0D
        Dim lstrMens = String.Empty
        Dim lstrNroFac = ClsPanorama.FstrNumeroDcto(aobjEstaCta.ObjPrefijoFac_EstadoStr.ObjValorPro,
                aobjEstaCta.ObjIdFactura_EstadoEnt.ObjValorPro)
        For Each lobjFactEsta As ClsFacturaEstado In aobjEstaCta.ColFacturasEstado
            ldecDedudaCap += lobjFactEsta.ObjDeudaCap_ItFacEstDec.ObjValorPro
            ldecDeudaInt += lobjFactEsta.ObjDeudaIntMora_ItFacEstDec.ObjValorPro
        Next
        If ldecDedudaCap <> aobjEstaCta.ObjDeudaCapitalDec.ObjValorPro Then
            lstrMens = "Estado de Cuenta Nro. " & aobjEstaCta.ObjIdEstadoCuentaEnt.ToString &
                    " de la factura " & lstrNroFac & " tiene diferencias en la Deuda de Capital!"
        End If
        If ldecDeudaInt <> aobjEstaCta.ObjDeudaIntMoraDec.ObjValorPro Then
            lstrMens = "Estado de Cuenta Nro. " & aobjEstaCta.ObjIdEstadoCuentaEnt.ToString &
                    " de la factura " & lstrNroFac & " tiene diferencias en la Ceuda de Capital!"
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            SAdicioneErrorIntegridad(lstrMens)
        End If
    End Sub
    Private Sub SAdicioneErrorIntegridad(astLinea As String)
        Dim lswMapImpo As StreamWriter
        lswMapImpo = ClsPanorama.FswStreamWriterAppend(MstrArchivoIntegridad)
        lswMapImpo.WriteLine(astLinea)
        lswMapImpo.Close()
        MblnHayerror = True
    End Sub
    Private Sub SAdicioneFinIntegridad(astLinea As String)
        Dim lswMapImpo As StreamWriter
        lswMapImpo = ClsPanorama.FswStreamWriterAppend(MstrArchivoIntegridad)
        lswMapImpo.WriteLine(astLinea)
        lswMapImpo.Close()
    End Sub
    Private Sub SInicialiceArchivos()
        Dim lswMapImpo As StreamWriter
        lswMapImpo = ClsPanorama.FswStreamWriter(MstrArchivoIntegridad)
        lswMapImpo.Close()
    End Sub
#End Region
#Region "Fechas y Periodos"
    ''' <summary>
    ''' Devuelve el periodo resultante (yyyymm) de sumarle al periodo pasado en el argumento "astrPeriodoInicial"
    '''  la cantidad de periodos pasada en el argumento "aentCantidadPeriodos"
    ''' </summary>
    ''' <param name="astrPeriodoInicial"></param>
    ''' <param name="aentCantidadPeriodos"></param>
    ''' <returns>Un string formado por los cuatro digitos del año seguido por los dos digitos del mes</returns>
    ''' <remarks></remarks>
    Friend Shared Function FstrPeriodoFinal(astrPeriodoInicial As String,
            aentCantidadPeriodos As Integer) As String
        Dim lentIdAnoFin As Integer, lentIdPerFinal As Integer
        Dim lentIdAno As Integer = CType(astrPeriodoInicial.Substring(0, 4), Integer)
        Dim lentIdPeriodo As Integer = CType(astrPeriodoInicial.Substring(4, 2), Integer)
        Dim lentPeriodoIni As Integer = (lentIdAno * 12) + lentIdPeriodo
        Dim lentPeriodoFinal = lentPeriodoIni + aentCantidadPeriodos
        If lentPeriodoFinal Mod 12 = 0 Then
            If aentCantidadPeriodos > 0 Then
                lentIdAnoFin = lentIdAno + Int(aentCantidadPeriodos / 12)
                lentIdPerFinal = 12
            ElseIf aentCantidadPeriodos = 0 Then
                lentIdAnoFin = lentIdAno
                lentIdPerFinal = lentIdPeriodo
            Else
                lentIdAnoFin = lentIdAno - 1
                lentIdPerFinal = 12
            End If
        Else
            lentIdAnoFin = Fix(lentPeriodoFinal / 12)
            lentIdPerFinal = lentPeriodoFinal - (lentIdAnoFin * 12)
        End If
        Dim lstrIdPeriodoFinal As String = CStr(lentIdAnoFin) &
                Format(lentIdPerFinal, "0#")
        Return lstrIdPeriodoFinal
    End Function
    ''' <summary>
    ''' Devuelve un string de seis caracteres númericos compuesto por los cuatro caracteres del año y 
    ''' dos caracteres del mes de la fecha pasada en el argumento.
    ''' </summary>
    ''' <param name="adtmFecha"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Shared Function FstrPeriodoDeFecha(adtmFecha As Date) As String
        Return ClsPanorama.FstrPeriodo(adtmFecha)
    End Function
    ''' <summary>
    ''' Devuelve la cantidad de periodos entre dos periodos definidos por una cadena de seis caracteres
    ''' donde los dos primeros son el año y los dos últimos el mes
    ''' </summary>
    ''' <param name="astrPeriodoIni">Periodo Inicial</param>
    ''' <param name="astrPeriodoFin">Periodo Final</param>
    ''' <remarks></remarks>
    Friend Shared Function FentCantPeriodosEntrePeriodos(astrPeriodoIni As String,
            astrPeriodoFin As String) As Integer
        Dim lentAnoIni As Integer = CType(astrPeriodoIni.Substring(0, 4), Integer)
        Dim lentMesIni As Integer = CType(astrPeriodoIni.Substring(4, 2), Integer)
        Dim lentCantPerIni As Integer = (lentAnoIni * 12) + lentMesIni
        Dim lentAnoFin As Integer = CType(astrPeriodoFin.Substring(0, 4), Integer)
        Dim lentMesFin As Integer = CType(astrPeriodoFin.Substring(4, 2), Integer)
        Dim lentCantPerFin As Integer = (lentAnoFin * 12) + lentMesFin
        Return lentCantPerFin - lentCantPerIni
    End Function
    ''' <summary>
    ''' Devuelve el período anterior al período del día de hoy
    ''' </summary>
    ''' <returns></returns>
    Friend Shared Function FobjPeriodoAnterior()
        Dim lstrPeriodActual = GobjParametros.ObjAnoActual.StrIdPeriodoActual
        Dim lstrPeriodoAnt = FstrPeriodoFinal(lstrPeriodActual, -1)
        Dim lobjAno As ClsAno = GobjParametros.ColAnos(lstrPeriodoAnt.Substring(0, 4))
        Dim lobjPeriodo As ClsPeriodo = lobjAno.ColPeriodos(lstrPeriodoAnt.Substring(4))
        Return lobjPeriodo
    End Function
#End Region
#Region "Cuentas Contabilidad"
    ''' <summary>
    ''' Indica si la cuenta de contabilidad pasada en el argumento, además de existir es una Cuenta Terminal!
    ''' </summary>
    ''' <param name="astrIdCtaContabilidad">String que contiene la Id de la Cuenta de Contabilidad formateada.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Shared Function FblnEsValidaCtaContabilidad(astrIdCtaContabilidad As String) As Boolean
        Dim lblnExiste = ClsCarpeta.FblnExisteCuentaCont(astrIdCtaContabilidad)
        Return lblnExiste
    End Function
    Friend Shared Function FstrNombreCuentaCon(astrIdCuentaCont As String) As String
        If String.IsNullOrEmpty(astrIdCuentaCont) Then
            Return String.Empty
        Else
            Return ObjMiCarpeta.FstrNombreCuentaCont(astrIdCuentaCont)
        End If
    End Function
    ''' <summary>
    ''' Devuelve el Nombre del Bamco seguido del número de la cuenta correspondiente a la cuenta
    ''' de conrabilidad pasada en el argumento.
    ''' </summary>
    ''' <param name="astrIdCtaContable">Id de la cuenta de contabilidad ligada a la cuenta
    ''' Bancaria que se devuelve</param>
    Friend Shared Function FstrCuentaBanco(astrIdCtaContable As String) As String
        Dim lstrEntFra = String.Empty, lstrCompl = String.Empty
        Dim lstrCuenta = String.Empty, lstrBanco = String.Empty
        If Not String.IsNullOrEmpty(astrIdCtaContable) Then
            Dim ldtbCuentasIng = FdtbCuentasIngresos()
            Dim lstrFiltro = ClsIdCtaContabilidadStr.SstrNombreCampoBd & " = '" & astrIdCtaContable & "'"
            Dim ldrwBancos As DataRow() = ldtbCuentasIng.Select(lstrFiltro)
            If ldrwBancos.Length > 0 Then
                lstrBanco = ldrwBancos(0)(ClsNombreBancoStr.SstrNombreCampoBd)
                lstrCuenta = ldrwBancos(0)(ClsNumeroCuentaStr.SstrNombreCampoBd)
            End If
            If astrIdCtaContable <> GobjParametros.ObjIdCtaCajaStr.ObjValorPro AndAlso
astrIdCtaContable <> GobjParametros.ObjIdCtaIngPorIdentificarStr.ObjValorPro Then
                lstrCompl = " Cta. Nro. "
            End If
            lstrEntFra = lstrBanco + lstrCompl + lstrCuenta
        End If
        Return lstrEntFra
    End Function
    ''' <summary>
    ''' Indica si una cuenta contable esta siendo utilizada o no, o si ya fue utilizada
    ''' </summary>
    ''' <param name="astrIdCuenta">Id de la cuenta a evaluar</param>
    Friend Shared Function FblnCtaConEsEliminables(astrIdCuenta As String) As Boolean
        Dim lblnEstaUsada = FblnEstaUsadaEnNov(astrIdCuenta) OrElse
                FblnEstaEnCentroUtil(astrIdCuenta) OrElse
                FblnEstaUsadaPorServicio(astrIdCuenta) OrElse
                FblnEstaUsadaEnBancos(astrIdCuenta)
        Return Not lblnEstaUsada
    End Function
    Private Shared Function FblnEstaEnCentroUtil(astrIdCuenta As String) As Boolean
        Dim lblnEsta = False
        With GobjParametros
            lblnEsta = (.ObjIdCtaAnticiposRecibidosStr.ObjValorPro = astrIdCuenta)
            lblnEsta = lblnEsta OrElse (.ObjIdCtaCajaStr.ObjValorPro = astrIdCuenta)
            lblnEsta = lblnEsta OrElse (.ObjIdCtaDescuentosPPStr.ObjValorPro = astrIdCuenta)
            lblnEsta = lblnEsta OrElse (.ObjIdCtaIngPorIdentificarStr.ObjValorPro = astrIdCuenta)
            lblnEsta = lblnEsta OrElse (.ObjIdCtaIntMoraDbStr.ObjValorPro = astrIdCuenta)
            lblnEsta = lblnEsta OrElse (.ObjIdCtaReteFuenteStr.ObjValorPro = astrIdCuenta)
            lblnEsta = lblnEsta OrElse (.ObjIdCtaReteIcaStr.ObjValorPro = astrIdCuenta)
            lblnEsta = lblnEsta OrElse (.ObjIdCtaReteIvaStr.ObjValorPro = astrIdCuenta)
        End With
        Return lblnEsta
    End Function
    Private Shared Function FblnEstaUsadaPorServicio(astrIdCuenta As String)
        Dim lblnEstaUsada = False
        For Each lobjAno As ClsAno In GobjParametros.ColAnos
            lblnEstaUsada = FblnEstaUsadaEnAno(lobjAno, astrIdCuenta)
            If lblnEstaUsada Then Exit For
        Next
        If Not lblnEstaUsada Then
            For Each lobjServicio As ClsServicio In GobjParametros.ColServiciosPer
                With lobjServicio
                    lblnEstaUsada = (.ObjCodigoCuentaDevStr.ObjValorPro = astrIdCuenta)
                    lblnEstaUsada = lblnEstaUsada OrElse (.ObjCodigoCuentaCrStr.ObjValorPro = astrIdCuenta)
                    lblnEstaUsada = lblnEstaUsada OrElse (.ObjCodigoCuentaDbStr.ObjValorPro = astrIdCuenta)
                    lblnEstaUsada = lblnEstaUsada OrElse (.ObjCodigoCuentaIvaStr.ObjValorPro = astrIdCuenta)
                    lblnEstaUsada = lblnEstaUsada OrElse (.ObjCodigoCuentaMoraStr.ObjValorPro = astrIdCuenta)
                End With
                If lblnEstaUsada Then Exit For
            Next
        End If
        Return lblnEstaUsada
    End Function
    Private Shared Function FblnEstaUsadaEnAno(aobjAno As ClsAno, astrIdCuenta As String) As Boolean
        Dim lblnEstaUsada = False
        For Each lobjServicio As ClsServicio In aobjAno.ColServiciosAno
            With lobjServicio
                lblnEstaUsada = (.ObjCodigoCuentaDevStr.ObjValorPro = astrIdCuenta)
                lblnEstaUsada = lblnEstaUsada OrElse (.ObjCodigoCuentaCrStr.ObjValorPro = astrIdCuenta)
                lblnEstaUsada = lblnEstaUsada OrElse (.ObjCodigoCuentaDbStr.ObjValorPro = astrIdCuenta)
                lblnEstaUsada = lblnEstaUsada OrElse (.ObjCodigoCuentaIvaStr.ObjValorPro = astrIdCuenta)
                lblnEstaUsada = lblnEstaUsada OrElse (.ObjCodigoCuentaMoraStr.ObjValorPro = astrIdCuenta)
            End With
            If lblnEstaUsada Then Exit For
        Next
        Return lblnEstaUsada
    End Function
    Private Shared Function FblnEstaUsadaEnBancos(astrIdCuenta As String) As Boolean
        Dim lblnEstaUsada = False
        Dim lcolCuentasBanco = GobjParametros.FcolCuentasBanco
        For Each lobjCuentaBanco As ClsCuentaBanco In lcolCuentasBanco
            lblnEstaUsada = lblnEstaUsada OrElse
                    (lobjCuentaBanco.ObjIdCtaContabilidadStr.ObjValorPro = astrIdCuenta)
        Next
        Return lblnEstaUsada
    End Function
    Private Shared Function FblnEstaUsadaEnNov(astrIdCuenta As String) As Boolean
        Dim lstrCamposSelect = {"COUNT(" & ClsIdNovedadShr.SstrNombreCampoBd & ")"}
        Dim lstrFiltro = ClsIdCuentaCr_NovStr.SstrNombreCampoBd & " = '" & astrIdCuenta & "' OR " &
                ClsIdCuentaDb_NovStr.SstrNombreCampoBd & " = '" & astrIdCuenta & "'"
        Dim ldtbRes = ClsPanorama.FdtbDataTable(ClsNovedad.SstrNombreTabla, lstrCamposSelect,
                {{}}, lstrFiltro)
        If ClsPanorama.FobjValorCampo(ldtbRes.Rows(0)(0), EnuTipoValor.enuInteger) > 0 Then
            Return True
        End If
        Return False
    End Function
#End Region
#Region "ImagenesQR"
    ''' <summary>
    ''' Devuelve un DataTable con la imagen del código QR utilizado para transferir 
    ''' a una cuenta bancaria de la Copropiedad
    ''' </summary>
    ''' <param name="astrPrefFac">Si es cero devuelve todos los registros de la copropiedad</param>
    ''' <remarks>En cada copropiedad solo puede haber un código QR. El Id de la imagen és el número
    ''' compuesto por el Id de la carpeta concatenado con el Id del centro de utilidad y concatenado
    ''' con el número 1</remarks>
    ''' <returns></returns>
    Friend Shared Function FdtbImagenBancoQR(aentIdBanco As Integer) As DataTable
        Dim lstrTabla = ClsImagen.SstrNombreTabla
        Dim lstrId As String = GshrIdCarpeta.ToString() & GshrIdCentroUtil.ToString()
        Dim lstrFiltro As String
        If aentIdBanco = 0 Then
            lstrFiltro = "IdImagen LIKE '" & lstrId & "%' AND IdTblCategoria = " &
                        EnuCategoriaImagenDef.enuDocumentos
        Else
            lstrId &= aentIdBanco.ToString()
            lstrFiltro = "IdImagen = " & lstrId & " AND IdTblCategoria = " &
                        EnuCategoriaImagenDef.enuDocumentos
        End If
        Dim ldtbQR = ClsPanorama.FdtbDataTable(lstrTabla, {"*"},
                {{"IdImagen", "ASC"}, {"IdTblCategoria", "ASC"}, {"Ordinal", "ASC"}},
                lstrFiltro)
        Return ldtbQR
    End Function
#End Region
#Region "Envio eMails"
    Friend Shared Function FstrArchivoPdfDcto(astrPrefDcto As String,
            aentIdDcto As Integer, aenuTipoDoc As EnuTipoDocOri) As String
        Dim lstbNomArch As New StringBuilder
        Dim lstrArch = FstrNomArchPdfDcto(astrPrefDcto, aentIdDcto, aenuTipoDoc)
        With lstbNomArch
            .Clear.Append(GstrTrayEmails).Append(lstrArch)
        End With
        Return lstbNomArch.ToString
    End Function
    Friend Shared Function FstrNomArchPdfDcto(astrPrefDcto As String,
            aentIdDcto As Integer, aenuTipoDoc As EnuTipoDocOri) As String
        Dim lstbArchPdf As New StringBuilder
        Dim lstrExtension = ".pdf", lstrDoc = String.Empty
        Dim lstrNroDoc = ClsPanorama.FstrNumeroDcto(astrPrefDcto, aentIdDcto)
        Select Case aenuTipoDoc
            Case EnuTipoDocOri.EnuFactura
                lstrDoc = "\Factura_"
            Case EnuTipoDocOri.EnuReciboCaja
                lstrDoc = "\RecCaja_"
            Case EnuTipoDocOri.EnuNotaCr
                lstrDoc = "\NotaCr_"
            Case EnuTipoDocOri.EnuNotaDb
                lstrDoc = "\NotaDb_"
            Case EnuTipoDocOri.EnuNotaCon
                lstrDoc = "\NotaCon_"
            Case EnuTipoDocOri.EnuNotaRevCr
                lstrDoc = "\NotaRCr_"
        End Select
        With lstbArchPdf
            .Clear.Append(lstrDoc)
            .Append(lstrNroDoc).Append(lstrExtension)
        End With
        Return lstbArchPdf.ToString
    End Function
    Friend Shared Function FstrNomArchEstadoCtaPreAgr(astrIdPredAgr As String)
        Dim lstrFecha = Today.Year.ToString() & Format(Today.Month, "0#") &
                Format(Today.Day, "0#")
        Dim lstrNomArch = GstrTrayEmails & "\" & lstrFecha & "_EstadoCuenta_" & astrIdPredAgr &
            ".pdf"
        Return lstrNomArch
    End Function
#End Region
#Region "Exporta Facturas PDF"
    Friend Shared Function FblnHayFacs(ablnSoloAutomaticas As Boolean) As Boolean
        Dim lblnHay = False
        Dim ldtmFechaIni = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
        Dim lstrFecha = ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaIni)
        lstrFecha = "'" & lstrFecha & "%'"
        If ldtmFechaIni <> GCDTMFECHANULA Then
            Dim lstrTabla = ClsFactura.SstrNombreTabla
            Dim lstrCamposSelect = {"COUNT(" & ClsIdFacturaEnt.SstrNombreCampoBd & ")"}
            Dim lstrFiltro = StrFiltroUbicacion & " AND " & ClsFechaFacturaDtm.SstrNombreCampoBd &
" >= " & lstrFecha
            If ablnSoloAutomaticas Then
                lstrFiltro &= " AND " & ClsIdModoFacturacionByt.SstrNombreCampoBd & " = " &
EnuModoFacturacionDef.EnuSistema
            End If
            Dim lstrOrden = {{"", ""}}
            Dim ldtbRes = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamposSelect, lstrOrden,
lstrFiltro)
            Dim lentCan = ClsPanorama.FobjValorCampo(ldtbRes.Rows(0)(0), EnuTipoValor.enuInteger)
            lblnHay = lentCan > 0
        End If
        Return lblnHay
    End Function
    Friend Shared Sub SCreeCarpetaFras()
        Dim lstrPeriodoActual = GobjParametros.ObjAnoActual.StrIdPeriodoActual
        GstrTrayFacturasPdf = GstrTrayReportes & "\" & lstrPeriodoActual & "_FacturasMes"
        If Not My.Computer.FileSystem.DirectoryExists(GstrTrayFacturasPdf) Then
            My.Computer.FileSystem.CreateDirectory(GstrTrayFacturasPdf)
        End If
    End Sub
    ''' <summary>
    ''' Devuelve el prefijo y el numero de las facturas del mes no exportdas y que son exportables
    ''' </summary>
    ''' <returns></returns>
    Friend Shared Function FdtbFacsMesExportar() As DataTable
        Dim ldtmFechaIni = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
        Dim lstrFechaIni = "'" & ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaIni) & "'"
        Dim lstrTabla = ClsFactura.SstrNombreTabla
        Dim lstrCamposSelect = {ClsPrefijo_FactStr.SstrNombreCampoBd,
                ClsIdFacturaEnt.SstrNombreCampoBd}
        Dim lstrOrden As String(,) = {{ClsPrefijo_FactStr.SstrNombreCampoBd, "ASC"},
                {ClsIdFacturaEnt.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " & ClsFechaFacturaDtm.SstrNombreCampoBd &
                " >= " & lstrFechaIni & " AND " & ClsPrefijo_FactStr.SstrNombreCampoBd &
                " <> '" & GCSTRPREFPREFACTURA & "'"
        If GobjParametros.BlnEFacAutorizado Then
            lstrFiltro &= " AND " & ClsIdEstadoEDocEnt.SstrNombreCampoBd & " >= " &
                    EnuEstadoEDoc.EnuRegi
        End If
        Dim ldtbFacs = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamposSelect, lstrOrden, lstrFiltro)
        Return ldtbFacs
    End Function
    ''' <summary>
    ''' Devuelve el prefijo y el numero de las facturas del mes no enviadas por email y
    ''' cuando está la facturación electrónica habilitada, las que estan en estado >= 4
    ''' </summary>
    ''' <returns></returns>
    Friend Shared Function FdtbFacsMesEnviarEmail() As DataTable
        Dim ldtmFechaIni = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
        Dim lstrFechaIni = "'" & ClsPanoramaDat.FstrFechaNormalizada(ldtmFechaIni) & "'"
        Dim lstrTabla = ClsFactura.SstrNombreTabla
        Dim lstrCamposSelect = {ClsPrefijo_FactStr.SstrNombreCampoBd,
ClsIdFacturaEnt.SstrNombreCampoBd}
        Dim lstrOrden As String(,) = {{ClsPrefijo_FactStr.SstrNombreCampoBd, "ASC"},
{ClsIdFacturaEnt.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " & ClsFechaFacturaDtm.SstrNombreCampoBd &
                " >= " & lstrFechaIni & " AND " & ClsEnviadaMailBln.SstrNombreCampoBd & " = False" &
                " AND " & ClsIdModoFacturacionByt.SstrNombreCampoBd & " = " &
                EnuModoFacturacionDef.EnuSistema & " AND " &
                ClsPrefijo_FactStr.SstrNombreCampoBd & " <> '" & GCSTRPREFPREFACTURA & "'"
        If GobjParametros.BlnEFacAutorizado Then
            lstrFiltro &= " AND " & ClsIdEstadoEDocEnt.SstrNombreCampoBd & " >= " &
                   EnuEstadoEDoc.EnuRegi
        End If
        Dim ldtbFacs = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamposSelect, lstrOrden, lstrFiltro)
        Return ldtbFacs
    End Function
#End Region
#Region "Exporta Facturas individuales a PDF"
    ''' <summary>
    ''' Devuelve el nombre del archivo PDF de la factura incluyendo la trayectoria completa
    ''' </summary>
    ''' <param name="astrPrefFac"></param>
    ''' <param name="aentIdFact"></param>
    ''' <returns></returns>
    Friend Shared Function FstrNombreArchFactura(astrPrefFac As String, aentIdFact As Integer) As String
        Dim lstrPeriodoActual = GobjParametros.ObjAnoActual.StrIdPeriodoActual
        Dim lstrCarpetaFacturasMes = GstrTrayReportes & "\" & lstrPeriodoActual & "_FacturasMes\"
        Dim lstrArchFac = FstrArchFactura(astrPrefFac, aentIdFact)
        Dim lstrNomArchFac = lstrCarpetaFacturasMes + lstrArchFac
        Return lstrNomArchFac
    End Function
    ''' <summary>
    ''' Devuelve el Nombre del Archivo PDF de la factura sin incluir la trayectoria
    ''' </summary>
    ''' <param name="astrPrefFac"></param>
    ''' <param name="aentIdFact"></param>
    ''' <returns></returns>
    Friend Shared Function FstrArchFactura(astrPrefFac As String, aentIdFact As Integer) As String
        Dim lstrArchFac = "Fac-"
        Dim lstrPredioAgr As String
        Dim lobjFactura As New ClsFactura()
        Dim lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, astrPrefFac, aentIdFact}
        lobjFactura.SAbra(lobjValorLlave)
        If Not lobjFactura.BlnExiste Then
            Throw New PanLException("Error, Factura en proceso exportacion no existe!")
        End If
        If String.IsNullOrEmpty(lobjFactura.ObjIdPredioAgrupador_FacStr.ObjValorPro) Then
            lstrArchFac += lobjFactura.ObjIdCliente_FactDbl.ToString
        Else
            lstrPredioAgr = lobjFactura.ObjIdPredioAgrupador_FacStr.ObjValorPro
            If lstrPredioAgr.Contains("/") Then
                lstrPredioAgr = lstrPredioAgr.Replace("/", "-")
            ElseIf lstrPredioAgr.Contains("*") Then
                lstrPredioAgr = lstrPredioAgr.Replace("*", "-")
            ElseIf lstrPredioAgr.Contains(Chr(34)) Then
                lstrPredioAgr = lstrPredioAgr.Replace(Chr(34), "-")
            End If
            lstrArchFac += lstrPredioAgr
        End If
        lstrArchFac += "_" & ClsPanorama.FstrNumeroDcto(astrPrefFac, aentIdFact) & ".Pdf"
        Return lstrArchFac
    End Function
#End Region
#Region "Exporta Recibos de Caja individuales a PDF"
    Friend Shared Sub SCreeCarpetaRec()
        GstrTrayRecibosCajaPdf = GstrTrayReportes & "\" & "Recibos de Caja"
        If Not My.Computer.FileSystem.DirectoryExists(GstrTrayRecibosCajaPdf) Then
            My.Computer.FileSystem.CreateDirectory(GstrTrayRecibosCajaPdf)
        End If
    End Sub
    Friend Shared Function FstrNombreArchRecibo(astrPrefRec As String, aentIdRec As Integer) As String
        Dim lstrNomArchRec = "Rec-"
        Dim lstrPredioAgr As String
        Dim lobjRecCaja As New ClsReciboCaja()
        Dim lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, astrPrefRec, aentIdRec}
        lobjRecCaja.SAbra(lobjValorLlave)
        If Not lobjRecCaja.BlnExiste Then
            Throw New PanLException("Error, Recibo de Caja en proceso exportacion no existe!")
        End If
        If String.IsNullOrEmpty(lobjRecCaja.ObjIdPredioAgrupador_RecStr.ObjValorPro) Then
            lstrNomArchRec += lobjRecCaja.ObjIdCliente_RecDbl.ToString
        Else
            lstrPredioAgr = lobjRecCaja.ObjIdPredioAgrupador_RecStr.ObjValorPro
            If lstrPredioAgr.Contains("/") Then
                lstrPredioAgr = lstrPredioAgr.Replace("/", "-")
            ElseIf lstrPredioAgr.Contains("*") Then
                lstrPredioAgr = lstrPredioAgr.Replace("*", "-")
            ElseIf lstrPredioAgr.Contains(Chr(34)) Then
                lstrPredioAgr = lstrPredioAgr.Replace(Chr(34), "-")
            End If
            If lstrPredioAgr.StartsWith(",") Then
                lstrPredioAgr = lstrPredioAgr.Substring(1)
            End If
            lstrNomArchRec += lstrPredioAgr
        End If
        lstrNomArchRec += "_" & ClsPanorama.FstrNumeroDcto(astrPrefRec, aentIdRec) & ".Pdf"
        lstrNomArchRec = GstrTrayEmails & "\" & lstrNomArchRec
        Return lstrNomArchRec
    End Function
#End Region
#Region "Copia de Seguridad"
    Friend Shared Function FstrNombreArchivoCopia(ablnCierreMes As Boolean) As String
        Dim lstrArchivo = String.Empty
        Dim lstrTrayCopia = ClsAdministrador.FobjAppActual.StrTrayCopiaSeguridad
        Dim lstrPrefijo = ClsPanorama.FstrFechayyyymmdd(Date.Today) & "_"
        If ablnCierreMes Then
            lstrPrefijo &= "CM_"
        End If
        lstrPrefijo &= "C" & GshrIdCarpeta.ToString & "_"
        lstrPrefijo += My.Resources.NombreBk
        If Not My.Computer.FileSystem.DirectoryExists(lstrTrayCopia) Then
            My.Computer.FileSystem.CreateDirectory(lstrTrayCopia)
            lstrArchivo = lstrTrayCopia & "\" & lstrPrefijo & "_01.zip"
        Else
            Dim lstrSufijo = String.Empty, i = 0
            Do While True
                i += 1
                If i > 99 Then
                    Throw New PanLException("Se alcanzo la cantidad máxima de copias en el día. " &
                            "La copia no se efectuó!")
                    Exit Do
                End If
                lstrSufijo = Format(i, "0#")
                lstrArchivo = lstrTrayCopia & "\" & lstrPrefijo & "_" & lstrSufijo & ".zip"
                If Not My.Computer.FileSystem.FileExists(lstrArchivo) Then
                    lstrArchivo = lstrArchivo.Replace("zip", "sql")
                    Exit Do
                End If
            Loop
        End If
        Return lstrArchivo
    End Function
#End Region
#End Region
#Region "Procedimientos y funciones Estado de la Aplicacion"
    Friend Shared Function FblnHayItemsPorFacturar() As Boolean
        Dim lblnHay = False
        Dim lobjServicio As ClsServicio
        Dim ldtbSerPorFac = FdtbServiciosPorFacturar(EnuDestiItemProgramaFact.EnuTodos)
        For Each ldrwSer As DataRow In ldtbSerPorFac.Rows
            Dim lshrIdAno = ClsPanorama.FobjValorCampo(ldrwSer(
                    ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd), EnuTipoValor.enuShort)
            Dim lshrIdServicio = ClsPanorama.FobjValorCampo(ldrwSer(
                    ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd), EnuTipoValor.enuShort)
            Dim lstrKey = lshrIdAno.ToString & "," & lshrIdServicio.ToString
            If lshrIdAno = 0 Then
                lobjServicio = GobjParametros.ColServiciosPer(lstrKey)
            Else
                lobjServicio = GobjParametros.ObjAnoActual.ColServiciosAno(lstrKey)
            End If
            lblnHay = lobjServicio.DtmFechaFacturacionPeriodoActual <= Date.Today
            If lblnHay Then Exit For
        Next
        Return lblnHay
    End Function
    Friend Shared Function FdtbServiciosPorFacturar(aenuDestItem As EnuDestiItemProgramaFact) _
            As DataTable
        Dim lstrIdAno = GobjParametros.ObjAnoActual.ObjIdAnoShr.ToString
        Dim lstrIdPeriActual = GobjParametros.ObjAnoActual.StrIdPeriodoActual.Substring(4, 2)
        Dim lstrPeriodoActual = GobjParametros.ObjAnoActual.StrIdPeriodoActual
        Dim lstrTabla = ClsItemProgramaFact.SstrNombreTabla
        Dim lstrCampSele = {"DISTINCT " & ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd,
                ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd}
        Dim lstrFiltro = StrFiltroUbicacion & " AND " &
                ClsSaldo_ItemProgramaFactDec.SstrNombreCampoBd & " / " &
                ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd & " > 0 AND (" &
                ClsCantidadPeriodosShr.SstrNombreCampoBd & " - (" &
                ClsSaldo_ItemProgramaFactDec.SstrNombreCampoBd & " / " &
                ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd & ")) + (SUBSTR(" &
                ClsPeriodoIni_ItemProgStr.SstrNombreCampoBd & ", 1, 4) * 12 + SUBSTR(" &
                ClsPeriodoIni_ItemProgStr.SstrNombreCampoBd & ", 5, 2) - 1) < (" &
                lstrIdAno & " * 12) + " & lstrIdPeriActual & " AND " &
                ClsPeriodoIni_ItemProgStr.SstrNombreCampoBd & " <= '" & lstrPeriodoActual & "'"
        If aenuDestItem = EnuDestiItemProgramaFact.EnuCliente Then
            lstrFiltro &= " AND " & ClsIdPredio_ItemProgramaFactStr.SstrNombreCampoBd & " <> ''"
        ElseIf aenuDestItem = EnuDestiItemProgramaFact.EnuPredio Then
            lstrFiltro &= " AND " & ClsIdCliente_ItemProgramaFactDbl.SstrNombreCampoBd & " <> 0"
        End If
        Dim lstrOrden = {{ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd, "DESC"},
                {ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd, "ASC"}}
        Dim ldtbServPorFact = ClsPanorama.FdtbDataTable(lstrTabla, lstrCampSele,
                lstrOrden, lstrFiltro)
        Return ldtbServPorFact
    End Function
    Friend Shared Function FblnHayPrefacturas() As Boolean
        Dim lblnHayFac = False
        DtmFechaFacturasAReversar = GCDTMFECHANULA
        Dim lentCantAProcesar As Integer
        Dim lstrFiltro = StrFiltroUbicacion & " AND " & ClsEsPreFacturaBln.SstrNombreCampoBd &
                " = True"
        Dim ldrwPrefacturas As DataRow() = ClsPanorama.FdrwDataRow(ClsFactura.SstrNombreTabla,
                {"COUNT(" & ClsIdFacturaEnt.SstrNombreCampoBd & ")",
                ClsFechaFacturaDtm.SstrNombreCampoBd}, Nothing, lstrFiltro)
        lentCantAProcesar = ClsPanorama.FobjValorCampo(
                ldrwPrefacturas(0)(ClsIdFacturaEnt.SstrNombreCampoBd), EnuTipoValor.EnuInteger)
        If lentCantAProcesar > 0 Then
            lblnHayFac = True
            DtmFechaFacturasAReversar = ClsPanorama.FobjValorCampo(
                    ldrwPrefacturas(0)(ClsFechaFacturaDtm.SstrNombreCampoBd), EnuTipoValor.EnuDate)
        End If
        Return lblnHayFac
    End Function
    Friend Shared Function FblnCrearAno() As Boolean
        Dim lblnCrearAno = True
        If GobjParametros.ObjAnoActual IsNot Nothing Then
            lblnCrearAno = GobjParametros.ObjAnoActual.ObjEstaCerradoAnoBln.ObjValorPro
        End If
        Return lblnCrearAno
    End Function
    ''' <summary>
    ''' Devuelve un string que indica los Id de los clientes con el email mal configurado.
    ''' </summary>
    ''' <returns></returns>
    Friend Shared Function FblnEmailOk(ByRef astrMens As String) As Boolean
        Dim lstrIdClientes = String.Empty, lblnEmailOk As Boolean
        Dim lstrTablaPri = ClsCliente.SstrNombreTabla
        Dim lstrTablaSec = ClsTercero.SstrNombreTabla
        Dim lstrCamSelPri = {ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCamSelSec = Array.Empty(Of String)
        Dim lstrCamPriRel = {ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCamSecRel = {ClsIdTerceroDbl.SstrNombreCampoBd}
        Dim lstrOrden(,) = {{"", ""}}
        Dim lstrFiltro = StrFiltroUbicacion_Pri & " AND " &
                ClsRecibeDocsPorEmailBln.SstrNombreCampoBd & " = TRUE " & " AND " &
                ClsEmailStr.SstrNombreCampoBd & " = ''"
        Dim ldtbCompr = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCamSelPri,
                lstrTablaSec, lstrCamSelSec, lstrCamPriRel, lstrCamSecRel, lstrOrden,
                lstrFiltro, Array.Empty(Of String), True)
        lblnEmailOk = ldtbCompr.Rows.Count = 0
        If Not lblnEmailOk Then
            For Each ldrwCli As DataRow In ldtbCompr.Rows
                lstrIdClientes &= ClsPanorama.FobjValorCampo(ldrwCli(0), EnuTipoValor.enuString) &
", "
            Next
        End If
        If Not String.IsNullOrEmpty(lstrIdClientes) Then
            lstrIdClientes = lstrIdClientes.Substring(0, lstrIdClientes.Length - 2)
            If lstrIdClientes.Contains(",") Then
                astrMens = "Los Clientes " & lstrIdClientes & " tienen desconfigurado el Email!"
            Else
                astrMens = "El Cliente " & lstrIdClientes & " tiene desconfigurado el Email!"
            End If
        End If
        Return lblnEmailOk
    End Function
    Friend Shared Function FblnDocPorProcesarEFac() As Boolean
        If GobjParametros.BlnEFacAutorizado Then
            Dim lentCanDoc As Integer
            Dim lstrCamposSel = {"COUNT(" & ClsIdFacturaEnt.SstrNombreCampoBd & ")"}
            Dim lstrFiltro = StrFiltroUbicacion & " AND " &
                    ClsIdEstadoEDocEnt.SstrNombreCampoBd & " < " & EnuEstadoEDoc.EnuEnviada
            Dim ldtb = ClsPanorama.FdtbDataTable(ClsFactura.SstrNombreTabla, lstrCamposSel, Nothing,
                    lstrFiltro)
            lentCanDoc = ClsPanorama.FobjValorCampo(ldtb.Rows(0)(0), EnuTipoValor.enuInteger)
            If lentCanDoc > 0 Then
                Return True
            End If
            lstrCamposSel = {"COUNT(" & ClsIdNotaDbEnt.SstrNombreCampoBd & ")"}
            ldtb = ClsPanorama.FdtbDataTable(ClsNotaDb.SstrNombreTabla, lstrCamposSel, Nothing,
                    lstrFiltro)
            lentCanDoc = ClsPanorama.FobjValorCampo(ldtb.Rows(0)(0), EnuTipoValor.enuInteger)
            If lentCanDoc > 0 Then
                Return True
            End If
            lstrCamposSel = {"COUNT(" & ClsIdNotaCrEnt.SstrNombreCampoBd & ")"}
            ldtb = ClsPanorama.FdtbDataTable(ClsNotaCr.SstrNombreTabla, lstrCamposSel, Nothing,
                    lstrFiltro)
            lentCanDoc = ClsPanorama.FobjValorCampo(ldtb.Rows(0)(0), EnuTipoValor.enuInteger)
            If lentCanDoc > 0 Then
                Return True
            End If
            lstrCamposSel = {"COUNT(" & ClsIdNotaReversaCrEnt.SstrNombreCampoBd & ")"}
            ldtb = ClsPanorama.FdtbDataTable(ClsNotaReversionCr.SstrNombreTabla, lstrCamposSel, Nothing,
                    lstrFiltro)
            lentCanDoc = ClsPanorama.FobjValorCampo(ldtb.Rows(0)(0), EnuTipoValor.enuInteger)
            If lentCanDoc > 0 Then
                Return True
            End If
            lstrCamposSel = {"COUNT(" & ClsIdNotaConEnt.SstrNombreCampoBd & ")"}
            ldtb = ClsPanorama.FdtbDataTable(ClsNotaCon.SstrNombreTabla, lstrCamposSel, Nothing,
                    lstrFiltro)
            lentCanDoc = ClsPanorama.FobjValorCampo(ldtb.Rows(0)(0), EnuTipoValor.enuInteger)
            If lentCanDoc > 0 Then
                Return True
            End If
        End If
        Return False
    End Function
    ''' <summary>
    ''' Cuando el año esta parametrizado para que a cada servicio de cuota de administración 
    ''' contribuya un solo módulo, se verifica que no exista un sector contribuyendo a más 
    ''' de un modulo o sea a más de un servicio de administración
    ''' </summary>
    ''' <returns></returns>
    Friend Shared Function FblnHaySectContriAdminMasDeUnaVes() As Boolean
        Dim lblnHaySecRep = False, lshrIdSectSig As Short, lshrIdSectAnt As Short, i = 1
        Dim lstrTabla = ClsSectorModulo.SstrNombreTabla
        Dim lstrCamposSelect = {ClsIdSector_SectorModuloShr.SstrNombreCampoBd}
        Dim lstrFiltro = StrFiltroUbicacion
        If GobjParametros.ObjAnoActual IsNot Nothing Then
            lstrFiltro &= " AND ("
            For Each lobjModuContr As ClsModuloContribucion In GobjParametros.ColModulos
                If lobjModuContr.ObjContribuyeCuotaAdminBln.ObjValorPro Then
                    lstrFiltro &= ClsIdModulo_SectorModuloShr.SstrNombreCampoBd & " = " &
lobjModuContr.ObjIdModuloShr.ObjValorPro & " OR "
                End If
            Next
            lstrFiltro = lstrFiltro.Substring(0, lstrFiltro.Length - 4) & ")"
        End If
        Dim lstrOrden As String(,) = {{ClsIdSector_SectorModuloShr.SstrNombreCampoBd, "ASC"}}
        Dim ldtbSec = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamposSelect, lstrOrden, lstrFiltro,
False, Array.Empty(Of String))
        Do While i <= ldtbSec.Rows.Count - 1
            lshrIdSectAnt = ClsPanorama.FobjValorCampo(ldtbSec.Rows(i - 1)(
                    ClsIdSector_SectorModuloShr.SstrNombreCampoBd), EnuTipoValor.EnuShort)
            lshrIdSectSig = ClsPanorama.FobjValorCampo(ldtbSec.Rows(i)(
                    ClsIdSector_SectorModuloShr.SstrNombreCampoBd), EnuTipoValor.EnuShort)
            lblnHaySecRep = lshrIdSectSig = lshrIdSectAnt
            If lblnHaySecRep Then
                Exit Do
            End If
            i += 1
        Loop
        Return lblnHaySecRep
    End Function
    Friend Shared Function FblnSectYaContriAdmin(ashrIdSector As Short) As Boolean
        Dim lblnSecExis As Boolean
        Dim lstrTabla = ClsSectorModulo.SstrNombreTabla
        Dim lstrCamposSelect = {ClsIdSector_SectorModuloShr.SstrNombreCampoBd}
        Dim lstrFiltro = StrFiltroUbicacion
        If GobjParametros.ObjAnoActual IsNot Nothing Then
            lstrFiltro &= " AND ("
            For Each lobjModuContr As ClsModuloContribucion In GobjParametros.ColModulos
                If lobjModuContr.ObjContribuyeCuotaAdminBln.ObjValorPro Then
                    lstrFiltro &= ClsIdModulo_SectorModuloShr.SstrNombreCampoBd & " = " &
lobjModuContr.ObjIdModuloShr.ObjValorPro & " OR "
                End If
            Next
            lstrFiltro = lstrFiltro.Substring(0, lstrFiltro.Length - 4) & ")"
        End If
        Dim lstrOrden As String(,) = {{ClsIdSector_SectorModuloShr.SstrNombreCampoBd, "ASC"}}
        Dim ldtbSec = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamposSelect, lstrOrden, lstrFiltro)
        lstrFiltro = ClsIdSector_SectorModuloServicioShr.SstrNombreCampoBd & " = " &
ashrIdSector.ToString()
        lblnSecExis = ldtbSec.Select(lstrFiltro).Length > 0
        Return lblnSecExis
    End Function
    Friend Shared Function FblnEstaServicioActivo(aobjServicio As ClsServicio) As Boolean
        Dim lblnEstaActivo As Boolean = aobjServicio.ObjGeneraProgramBln.ObjValorPro
        If lblnEstaActivo Then
            Dim lshrIdAno As Short = aobjServicio.ObjIdAno_ServicioShr.ObjValorPro
            Dim lshrIdServicio As Short = aobjServicio.ObjIdServicioShr.ObjValorPro
            Dim lstrTabla = ClsItemProgramaFact.SstrNombreTabla
            Dim lstrCampSel As String() = {"*"}
            Dim lstrOrden = {{"", ""}}
            Dim lstrFiltro = StrFiltroUbicacion & " AND " &
                    ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " = " & lshrIdAno & " AND " &
                    ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd & " = " & lshrIdServicio &
                    " AND " & ClsSaldo_ItemProgramaFactDec.SstrNombreCampoBd & " > 0"
            Dim ldtbRes = ClsPanorama.FdtbDataTable(lstrTabla, lstrCampSel, lstrOrden, lstrFiltro)
            lblnEstaActivo = ldtbRes.Rows.Count > 0
        End If
        Return lblnEstaActivo
    End Function
#End Region
#Region "Fijar Ventana"
    ''' <summary>
    ''' Para mantener la ventana siempre visible
    ''' </summary>
    ''' <remarks>No utilizamos el valor devuelto</remarks>
    <DllImport("user32.DLL")>
    Private Shared Sub SetWindowPos(hWnd As Integer, hWndInsertAfter As Integer,
            X As Integer, Y As Integer, cx As Integer, cy As Integer, wFlags As Integer)
    End Sub
    Public Shared Sub SiempreEncima(handle As Integer)
        SetWindowPos(handle, -1, 0, 0, 0, 0, wFlags)
    End Sub
#End Region
#Region "Dispose"
    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
    Protected Overridable Overloads Sub Dispose(ablnDisposing As Boolean)
        If Not MblnDisposed Then
            If ablnDisposing Then
                MblnDisposed = True
            End If
        End If
    End Sub
#End Region
End Class
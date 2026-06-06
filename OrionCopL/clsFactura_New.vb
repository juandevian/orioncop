Imports System.Drawing
Imports System.IO
Imports System.Windows.Media.Imaging
Imports ThoughtWorks.QRCode.Codec
Friend Class ClsFactura
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriFacturas"
    ' Estructura
    Private MstcIntMora As StcIntMoraFactura() = Nothing
    ' Variables de modulo
    Private McolItemsFactura As Collection = Nothing
    Private MdtbItemsFact As DataTable = Nothing
    Private MobjPredioAgrFac As ClsPredio = Nothing
    Private MdtbNovedadesFact As DataTable = Nothing
    Private MobjClienteFactura As ClsCliente = Nothing
    Private MobjEstadoCuenta As ClsEstadoCuenta = Nothing
    Private MdtbEstadoCuenta As DataTable = Nothing
    Private MstrNroFactura As String = String.Empty
    Private MstrNroFactItem As String = String.Empty
    Private MobjNotaCrAnulo As ClsNotaCr = Nothing
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia un objeto factura en modo único
    ''' </summary>
    Public Sub New()
        HobjPadre = Nothing
        ObjClienteFactura = Nothing
        HblnEsCreable = False
        HblnEsAnulable = False
        HblnEsSuprimible = False
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuUnico
        HcolTablas.Add(MCSTRNOMBRETABLA)
        HcolCamposSelect.Add({"*"})
    End Sub
    ''' <summary>
    ''' Instancia un objeto factura en modo navegable
    ''' </summary>
    Public Sub New(astrPref As String)
        HobjPadre = Nothing
        ObjClienteFactura = Nothing
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion
        lstrFiltro &= " AND " & ClsPrefijo_FactStr.SstrNombreCampoBd & " = '" & astrPref & "'"
        HcolFiltros.Add(lstrFiltro)
        Dim lstrCamposSelect = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsPrefijo_FactStr.SstrNombreCampoBd, ClsIdFacturaEnt.SstrNombreCampoBd}
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuNavegable
        HblnEsModificable = False
        HblnEsSuprimible = False
        HcolTablas.Add(MCSTRNOMBRETABLA)
        HcolCamposSelect.Add(lstrCamposSelect)
    End Sub
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Cliente al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwObjeto">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Public Sub New(aobjPadre As ClsCliente, adrwObjeto As DataRow)
        HobjPadre = aobjPadre
        ObjClienteFactura = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        '
        DrwRegistroActual = adrwObjeto
        If Not IsNothing(DrwRegistroActual) Then
            DtbTablaColeccion = DrwRegistroActual.Table
        End If
    End Sub
#End Region
#Region "Propiedades"
#Region "Propiedades indentificadoras"
    Protected Overrides ReadOnly Property HstrNombreTabla As String
        Get
            Return MCSTRNOMBRETABLA
        End Get
    End Property
    Friend Shared ReadOnly Property SstrNombreTabla As String
        Get
            Return MCSTRNOMBRETABLA
        End Get
    End Property
    Protected Overrides ReadOnly Property HenuIdClase As EnuIdClasesPanDef
        Get
            Return EnuIdClasesPanDef.enuFactura
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Factura"
        End Get
    End Property
    Friend Overrides ReadOnly Property HstrNombreObj As String
        Get
            Return Chr(34) & "Nro. " & StrNumeroFactura & Chr(34)
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjCreditos_FactDec As New ClsCreditos_FactDec(Me)
    Friend ReadOnly Property ObjCUDocStr As New ClsCUDocStr(Me)
    Friend ReadOnly Property ObjCUFEStr As New ClsCUFEStr(Me)
    Friend ReadOnly Property ObjDctoProntoPago_FacDec As New ClsDctoProntoPago_FacDec(Me)
    Friend ReadOnly Property ObjDebitos_FactDec As New ClsDebitos_FactDec(Me)
    Friend ReadOnly Property ObjEnviadaMailBln As New ClsEnviadaMailBln(Me)
    Friend ReadOnly Property ObjEsPreFacturaBln As New ClsEsPreFacturaBln(Me)
    Friend ReadOnly Property ObjFechaAnulacion_FactDtm As New ClsFechaAnulacion_FactDtm(Me)
    Friend ReadOnly Property ObjFechaCancelacion_FactDtm As New ClsFechaCancelacion_FactDtm(Me)
    Friend ReadOnly Property ObjFechaDctoProntoPagoDtm As New ClsFechaDctoProntoPagoDtm(Me)
    Friend ReadOnly Property ObjFechaEmisionEFacStr As New ClsFechaEmisionEFacStr(Me)
    Friend ReadOnly Property ObjFechaFacturaDtm As New ClsFechaFacturaDtm(Me)
    Friend ReadOnly Property ObjFechaGraciaDtm As New ClsFechaGraciaDtm(Me)
    Friend ReadOnly Property ObjFechaVencimientoDtm As New ClsFechaVencimientoDtm(Me)
    Friend ReadOnly Property ObjIdCarpeta_FactShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_FactShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdEstadoEDocEnt As New ClsIdEstadoEDocEnt(Me)
    Friend ReadOnly Property ObjIdFacturaEnt As New ClsIdFacturaEnt(Me)
    Friend ReadOnly Property ObjIdFormaPagoByt As New ClsIdFormaPagoByt(Me)
    Friend ReadOnly Property ObjIdInformeCont_FacEnt As New ClsIdInformeCont_FacEnt(Me)
    Friend ReadOnly Property ObjIdMedioPagoByt As New ClsIdMedioPagoByt(Me)
    Friend ReadOnly Property ObjIdModoFacturacionByt As New ClsIdModoFacturacionByt(Me)
    Friend ReadOnly Property ObjIdPredioAgrupador_FacStr As New ClsIdPredioAgrupador_FacStr(Me)
    Friend ReadOnly Property ObjIdCliente_FactDbl As New ClsIdCliente_FactDbl(Me)
    Friend ReadOnly Property ObjIdUsuario_FactStr As New ClsIdUsuarioStr(Me)
    Friend ReadOnly Property ObjNumeroResolAutoStr As New ClsNumeroResolAutoStr(Me)
    Friend ReadOnly Property ObjPieFacturaDos_FactStr As New ClsPieFacturaDos_FactStr(Me)
    Friend ReadOnly Property ObjPieFacturaUno_FactStr As New ClsPieFacturaUno_FactStr(Me)
    Friend ReadOnly Property ObjPrefijo_FactStr As New ClsPrefijo_FactStr(Me)
    Friend ReadOnly Property ObjReferenciaPago_FacStr As New ClsReferenciaPago_FacStr(Me)
    Friend ReadOnly Property ObjValor_FactDec As New ClsValor_FactDec(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjAnuladoBln)
                HcolPropiedades.Add(ObjFechaCreacionDtm)
                HcolPropiedades.Add(ObjIdUsuarioAnuloStr)
                HcolPropiedades.Add(ObjOrigenInstanciaStr)
                HcolPropiedades.Add(ObjOrigenInstanciaAnuloStr)
                HcolPropiedades.Add(ObjCreditos_FactDec)
                HcolPropiedades.Add(ObjDctoProntoPago_FacDec)
                HcolPropiedades.Add(ObjDebitos_FactDec)
                HcolPropiedades.Add(ObjEsPreFacturaBln)
                HcolPropiedades.Add(ObjIdEstadoEDocEnt)
                HcolPropiedades.Add(ObjFechaAnulacion_FactDtm)
                HcolPropiedades.Add(ObjFechaCancelacion_FactDtm)
                HcolPropiedades.Add(ObjFechaDctoProntoPagoDtm)
                HcolPropiedades.Add(ObjFechaEmisionEFacStr)
                HcolPropiedades.Add(ObjFechaFacturaDtm)
                HcolPropiedades.Add(ObjFechaGraciaDtm)
                HcolPropiedades.Add(ObjFechaVencimientoDtm)
                HcolPropiedades.Add(ObjIdCarpeta_FactShr)
                HcolPropiedades.Add(ObjIdCentroUtil_FactShr)
                HcolPropiedades.Add(ObjIdFacturaEnt)
                HcolPropiedades.Add(ObjIdFormaPagoByt)
                HcolPropiedades.Add(ObjIdInformeCont_FacEnt)
                HcolPropiedades.Add(ObjIdMedioPagoByt)
                HcolPropiedades.Add(ObjIdModoFacturacionByt)
                HcolPropiedades.Add(ObjIdPredioAgrupador_FacStr)
                HcolPropiedades.Add(ObjIdCliente_FactDbl)
                HcolPropiedades.Add(ObjIdUsuario_FactStr)
                HcolPropiedades.Add(ObjPieFacturaDos_FactStr)
                HcolPropiedades.Add(ObjPieFacturaUno_FactStr)
                HcolPropiedades.Add(ObjPrefijo_FactStr)
                HcolPropiedades.Add(ObjValor_FactDec)
                HcolPropiedades.Add(ObjCUDocStr)
                HcolPropiedades.Add(ObjCUFEStr)
                HcolPropiedades.Add(ObjEnviadaMailBln)
                HcolPropiedades.Add(ObjNumeroResolAutoStr)
                HcolPropiedades.Add(ObjReferenciaPago_FacStr)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    Friend Property ObjClienteFactura As ClsCliente
        Set(value As ClsCliente)
            MobjClienteFactura = value
        End Set
        Get
            Dim lobjValorLlave As Object() = {ObjIdCarpeta_FactShr.ObjValorPro,
                ObjIdCentroUtil_FactShr.ObjValorPro, ObjIdCliente_FactDbl.ObjValorPro}
            MobjClienteFactura = New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
            MobjClienteFactura.SAbra(lobjValorLlave)
            Return MobjClienteFactura
        End Get
    End Property
    Friend ReadOnly Property ObjNotaCrAnulo As ClsNotaCr
        Get
            If IsNothing(MobjNotaCrAnulo) AndAlso ObjAnuladoBln.ObjValorPro Then
                Dim lobjItemFac As ClsItemFactura = ColItemsFactura("1")
                For Each lobjNov As ClsNovedad In lobjItemFac.ColNovedades
                    If lobjNov.ObjIdTipoDocOrigenByt.ObjValorPro = EnuTipoDocOri.EnuNotaCr Then
                        Dim lstrPrefNcr As String = lobjNov.ObjPrefijoDocOrigen_NovStr.ObjValorPro
                        Dim lentIdNcr As Integer = lobjNov.ObjIdDocOrigenEnt.ObjValorPro
                        MobjNotaCrAnulo = New ClsNotaCr()
                        MobjNotaCrAnulo.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefNcr, lentIdNcr})
                        If MobjNotaCrAnulo.ObjIdTipoNotaCrByt.ObjValorPro = EnuTipoNotaCrDef.EnuAnulaFac Then
                            Exit For
                        End If
                        MobjNotaCrAnulo = Nothing
                    End If
                Next
            End If
            Return MobjNotaCrAnulo
        End Get
    End Property
    ''' <summary>
    ''' Devuelve uns string compuesto por el prefijo de la factura y el id de la factura separados por un
    ''' guion. Si no existe el prefijo devuelve solo el id de la factura
    ''' </summary>
    ''' <value></value>
    Friend ReadOnly Property StrNumeroFactura As String
        Get
            Dim lstrNumeroFactura As String = ClsPanorama.FstrNumeroDcto(
                    ObjPrefijo_FactStr.ObjValorPro, ObjIdFacturaEnt.ObjValorPro)
            Return lstrNumeroFactura
        End Get
    End Property
    Friend ReadOnly Property ObjPredioAgrFactura As ClsPredio
        Get
            If IsNothing(MobjPredioAgrFac) Then
                If Not String.IsNullOrEmpty(ObjIdPredioAgrupador_FacStr.ToString) AndAlso
                        ObjIdPredioAgrupador_FacStr.ToString <> GCSTRSINPA Then
                    MobjPredioAgrFac = New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
                    Dim lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil,
                            ObjIdPredioAgrupador_FacStr.ObjValorPro}
                    MobjPredioAgrFac.SAbra(lobjValorLlave)
                    If Not MobjPredioAgrFac.BlnExiste Then
                        MobjPredioAgrFac = Nothing
                    End If
                End If
            End If
            Return MobjPredioAgrFac
        End Get
    End Property
    Friend ReadOnly Property StrNombrePredioAgr As String
        Get
            Dim lstrNombreredioAgr As String
            If ObjPredioAgrFactura Is Nothing Then
                lstrNombreredioAgr = "Sin Predio Agrupador"
            Else
                lstrNombreredioAgr = ObjPredioAgrFactura.ObjNombrePredioStr.ToString
            End If
            Return lstrNombreredioAgr
        End Get
    End Property
    ''' <summary>
    ''' Devuelve el valor total a pagar al momento de generar la factura, igual al valor de la 
    ''' factura mas la deuda pendiente según el objeto estado cuenta.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend ReadOnly Property DecTotalAPagar
        Get
            Dim ldecVlrAPagar As Decimal = ObjValor_FactDec.ObjValorPro
            If Not IsNothing(ObjEstadoCuenta) Then
                ldecVlrAPagar += ObjEstadoCuenta.ObjDeudaCapitalDec.ObjValorPro +
                        ObjEstadoCuenta.ObjDeudaIntMoraDec.ObjValorPro
            End If
            Return ldecVlrAPagar
        End Get
    End Property
    Friend ReadOnly Property EnuEstadoFactura As EnuEstadoFacturaDef
        Get
            Dim lenuEstado As EnuEstadoFacturaDef
            If DecDeuda = 0 Then
                If BlnExiste Then
                    If ObjAnuladoBln.ObjValorPro Then
                        lenuEstado = EnuEstadoFacturaDef.EnuAnulada
                    Else
                        lenuEstado = OrionCopL.EnuEstadoFacturaDef.EnuCancelada
                    End If
                Else
                    lenuEstado = EnuEstadoFacturaDef.None
                End If
            Else
                Dim ldtmFechaVenc As Date = ObjFechaVencimientoDtm.ObjValorPro
                Dim ldtmFechaGracia = ObjFechaGraciaDtm.ObjValorPro
                Dim ldtmHoy = Date.Today
                If ldtmHoy > ldtmFechaVenc Then
                    If ldtmHoy <= ldtmFechaGracia Then
                        lenuEstado = OrionCopL.EnuEstadoFacturaDef.EnuPeriodoGracia
                    Else
                        lenuEstado = OrionCopL.EnuEstadoFacturaDef.EnuVencida
                    End If
                Else
                    lenuEstado = OrionCopL.EnuEstadoFacturaDef.EnuNormal
                End If
            End If
            Return lenuEstado
        End Get
    End Property
    Friend ReadOnly Property DtmFechaPlazo As Date
        Get
            McolItemsFactura = ColItemsFactura
            Dim lobjItemFac As ClsItemFactura = McolItemsFactura(1)
            Dim lshrDiasGracia As Short = lobjItemFac.ShrDiasGracia
            Dim ldtmFechaVence As Date = ObjFechaVencimientoDtm.ObjValorPro
            Return ldtmFechaVence.AddDays(lshrDiasGracia)
        End Get
    End Property
    Friend ReadOnly Property BlnEnviarPorCorreo As Boolean
        Get
            Return ObjClienteFactura.ObjRecibeDocsPorEmailBln.ObjValorPro
        End Get
    End Property
    Friend ReadOnly Property BlnEsAdministracion As Boolean
        Get
            Dim lblnEs = False
            For Each lobjItemFac As ClsItemFactura In ColItemsFactura
                lblnEs = lobjItemFac.ObjServicio.BlnEsCuotaAdministracion
                If lblnEs Then
                    Exit For
                End If
            Next
            Return lblnEs
        End Get
    End Property
    Friend ReadOnly Property BlnFacturaPorServicio As Boolean
        Get
            Return If(ObjPredioAgrFactura IsNot Nothing,
                    ObjPredioAgrFactura.ObjFacturarPorServicio_PreBln.ObjValorPro,
                    ObjClienteFactura.ObjFactPorServicio_CliBln.ObjValorPro)
        End Get
    End Property
    Friend ReadOnly Property BlnFacturaAPropYPreAgr As Boolean
        Get
            Dim lobjItemFac As ClsItemFactura = ColItemsFactura(1)
            Return lobjItemFac.ObjServicio.ObjFactAPropYPreAgrBln.ObjValorPro
        End Get
    End Property
    Friend ReadOnly Property StrKeySerUnico As String
        Get
            Dim lstrKeySer = String.Empty
            If BlnFacturaPorServicio Then
                Dim lobjItemFac As ClsItemFactura = ColItemsFactura("1")
                lstrKeySer = lobjItemFac.ObjServicio.ObjIdAno_ServicioShr.ToString & "," &
                        lobjItemFac.ObjServicio.ObjIdServicioShr.ToString
            End If
            Return lstrKeySer
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SVacie()
        MyBase.SVacie()
        If MobjEstadoCuenta IsNot Nothing Then
            MobjEstadoCuenta.SVacie()
        End If
        MdtbItemsFact = Nothing
        McolItemsFactura = Nothing
        MstcIntMora = Nothing
        MdtbNovedadesFact = Nothing
        MobjEstadoCuenta = Nothing
        MdtbEstadoCuenta = Nothing
        MstrNroFactura = String.Empty
        MstrNroFactItem = String.Empty
        MobjClienteFactura = Nothing
        MobjNotaCrAnulo = Nothing
        MobjPredioAgrFac = Nothing
    End Sub
    Public Overrides Function FblnEsAnulable() As Boolean
        Dim lblnEsAnulable = BlnEsAnulable
        lblnEsAnulable = lblnEsAnulable AndAlso Not ObjAnuladoBln.ObjValorPro AndAlso
                ObjPrefijo_FactStr.ObjValorPro <> GCSTRPREFPREFACTURA
        If lblnEsAnulable Then
            If GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro Then
                lblnEsAnulable = (Date.Today <= GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo)
            End If
        End If
        If lblnEsAnulable Then
            If BlnEsFacEle Then
                lblnEsAnulable = BlnEstaRegEFac
            End If
        End If
        If lblnEsAnulable Then
            lblnEsAnulable = (FdecValorPagado() = 0)
        End If
        Return lblnEsAnulable
    End Function
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        Dim lblnNoHayError = False
        GobjPanDat.SControleProcesoObj(True)
        Try
            GobjPanDat.SInicialiceTransaccion()
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
                SEstablezcaDsctoPP()
                If Not ClsOrionCop.BlnProcesoEspecial Then
                    SEstablezcaFechas()
                End If
                If ObjIdModoFacturacionByt.ObjValorPro = EnuModoFacturacionDef.EnuContingencia Then
                    SCompleteFactCont()
                Else
                    SNumereObj()
                End If
                ClsPanorama.SActualiceCol(ColItemsFactura)
                SComplementeEFac()
                ObjFechaCreacionDtm.ObjValorPro = Date.Now
                MyBase.SActualice(ablnExigeRequeridos)
            Else
                ClsPanorama.SActualiceCol(ColItemsFactura)
                ObjValor_FactDec.SValide()
                ObjDebitos_FactDec.SValide()
                MyBase.SActualice(ablnExigeRequeridos)
            End If
            lblnNoHayError = True
        Catch ex As ErrorInesperadoPanLException
            Throw
        Catch ex As PanLException
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
    Protected Overrides Function SAnuleEnObj() As Boolean
        Dim lblnAnulado = FblnEsAnulable()
        If lblnAnulado Then
            ObjAnuladoBln.ObjValorPro = True
            ObjIdUsuarioAnuloStr.ObjValorPro = GstrIdUsuario
            ObjOrigenInstanciaAnuloStr.ObjValorPro = GstrOrigenActual
            If GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro Then
                ObjFechaAnulacion_FactDtm.ObjValorPro = Now
            Else
                If GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo < Date.Today Then
                    ObjFechaAnulacion_FactDtm.ObjValorPro =
                            GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
                Else
                    ObjFechaAnulacion_FactDtm.ObjValorPro = Now
                End If
            End If
            ' Items Factura
            For Each lobjItemFra As ClsItemFactura In ColItemsFactura
                lblnAnulado = lobjItemFra.SAnuleEnObj()
                If Not lblnAnulado Then Exit For
            Next
            If lblnAnulado Then
                ' Genera Nota Cr de anulación
                Dim lstrPref = GobjParametros.FstrPrefijoDoc(EnuIdDocumentoDef.EnuNotaCr)
                Dim lobjNotaCr As New ClsNotaCr(lstrPref)
                lobjNotaCr.SCreeObj(Nothing)
                lobjNotaCr.SAnuleFactura(Me)
                If lblnAnulado Then
                    ObjCreditos_FactDec.ObjValorPro = ObjDebitos_FactDec.ObjValorPro
                End If
                ' Reversar items programa facturación
                Dim lobjOrionCop As New ClsOrionCop(GCOBJREGISTRO, False)
                Dim larlNrosFact = lobjOrionCop.FarlFactsPropietarios(Me)
                If larlNrosFact(0) = StrNumeroFactura Then
                    If ObjIdModoFacturacionByt.ObjValorPro =
                            EnuModoFacturacionDef.EnuSistema Then
                        For Each lobjItemFra As ClsItemFactura In ColItemsFactura
                            lobjItemFra.SReverseItemProgFact(ObjIdCliente_FactDbl.ObjValorPro)
                        Next
                    End If
                End If
            End If
        End If
        Return lblnAnulado
    End Function
    Protected Overrides Sub SInicialiceObj()
        MyBase.SInicialiceObj()
        ObjAnuladoBln.ObjValorPro = False
        ObjPrefijo_FactStr.ObjValorPro =
                GobjParametros.FstrPrefijoDoc(EnuTipoDocOri.EnuFactura)
        ObjIdFacturaEnt.ObjValorPro = 0
        ObjFechaCreacionDtm.ObjValorPro = Date.Now
        ObjIdCarpeta_FactShr.ObjValorPro = GshrIdCarpeta
        ObjIdCentroUtil_FactShr.ObjValorPro = GshrIdCentroUtil
        ObjIdUsuario_FactStr.ObjValorPro = GstrIdUsuario
        ObjOrigenInstanciaStr.ObjValorPro = GstrOrigenActual
        ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuNoEDoc
        ObjCUFEStr.ObjValorPro = ""
    End Sub
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Dim lstrIdIbjeto = ObjIdFacturaEnt.ToString
            If ObjPrefijo_FactStr.ToString.Length > 0 Then
                lstrIdIbjeto = ObjPrefijo_FactStr.ObjValorPro & "-" & lstrIdIbjeto
            End If
            Return lstrIdIbjeto
        End Get
    End Property
#End Region
#Region "Procedimientos del objeto"
    Private Sub SNumereObj()
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
            If Not ClsOrionCop.BlnProcesoEspecial Then
                Dim lentIdFactura As Integer, lstrFiltroPre As String, lstrFiltroFac As String
                Dim lstrPrefijoPre As String = GCSTRPREFPREFACTURA
                Dim lstrPrefijoFac As String = GobjParametros.FstrPrefijoDoc(
                        EnuTipoDocOri.EnuFactura)
                If IsNothing(lstrPrefijoFac) Then lstrPrefijoFac = String.Empty
                lstrFiltroPre = ClsOrionCop.StrFiltroUbicacion & " AND " &
                        ClsPrefijo_FactStr.SstrNombreCampoBd & " = '" & lstrPrefijoPre & "'"
                lstrFiltroFac = ClsOrionCop.StrFiltroUbicacion & " AND " &
                        ClsPrefijo_FactStr.SstrNombreCampoBd & " = '" & lstrPrefijoFac & "'"
                If ObjEsPreFacturaBln.ObjValorPro Then
                    lentIdFactura = ClsPanorama.FobjUltimaIdNumericaObjeto(SstrNombreTabla,
                            ClsIdFacturaEnt.SstrNombreCampoBd, ObjIdFacturaEnt.EnuTipoValor,
                            lstrFiltroPre)
                    If lentIdFactura = 0 Then
                        lentIdFactura = ClsPanorama.FobjUltimaIdNumericaObjeto(SstrNombreTabla,
                                ClsIdFacturaEnt.SstrNombreCampoBd, ObjIdFacturaEnt.EnuTipoValor,
                                lstrFiltroFac)
                    End If
                Else
                    lentIdFactura = ClsPanorama.FobjUltimaIdNumericaObjeto(SstrNombreTabla,
                            ClsIdFacturaEnt.SstrNombreCampoBd, ObjIdFacturaEnt.EnuTipoValor,
                            lstrFiltroFac)
                End If
                If lentIdFactura < GobjParametros.FentNumeracionInicialDoc(
                        EnuTipoDocOri.EnuFactura) Then
                    lentIdFactura = GobjParametros.FentNumeracionInicialDoc(
                         EnuTipoDocOri.EnuFactura)
                End If
                lentIdFactura += 1
                If ObjEsPreFacturaBln.ObjValorPro Then
                    ObjPrefijo_FactStr.ObjValorPro = lstrPrefijoPre
                Else
                    ObjPrefijo_FactStr.ObjValorPro = lstrPrefijoFac
                End If
                ObjIdFacturaEnt.ObjValorPro = lentIdFactura
                ObjNumeroResolAutoStr.ObjValorPro =
                        GobjParametros.ObjNumeroResolFacturaStr.ObjValorPro
            End If
            SCompleteItemsFact()
        End If
    End Sub
    Private Sub SCompleteFactCont()
        ObjNumeroResolAutoStr.ObjValorPro = GobjParametros.ObjNumeroResolContiStr.ObjValorPro
        SCompleteItemsFact()
    End Sub
    Private Sub SComplementeEFac()
        If GobjParametros.ObjAutorizaEFacBln.ObjValorPro Then
            If ObjIdModoFacturacionByt.ObjValorPro = EnuModoFacturacionDef.EnuManual Then
                ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuNoReg
            End If
        End If
    End Sub
    Private Sub SCompleteItemsFact()
        Dim lobjItemFac As ClsItemFactura
        For i = 1 To ColItemsFactura.Count
            lobjItemFac = ColItemsFactura(i)
            lobjItemFac.ObjIdItemFacturaShr.ObjValorPro = i
            lobjItemFac.ObjPrefijo_ItemFactStr.ObjValorPro = ObjPrefijo_FactStr.ObjValorPro
            lobjItemFac.ObjIdFactura_ItemFactEnt.ObjValorPro = ObjIdFacturaEnt.ObjValorPro
            lobjItemFac.ObjEsPrefactura_ItemFactBln.ObjValorPro = ObjEsPreFacturaBln.ObjValorPro
        Next
    End Sub
    Private Sub SEstablezcaDsctoPP()
        Dim lblnTieneDsctoPP = FblnAplicaDsctoPP()
        If lblnTieneDsctoPP Then
            ObjDctoProntoPago_FacDec.ObjValorPro = FdecDsctoPPPosible()
            ObjFechaDctoProntoPagoDtm.ObjValorPro = FdtmFechaDsctoPP()
        Else
            ObjDctoProntoPago_FacDec.ObjValorPro = 0
            ObjFechaDctoProntoPagoDtm.ObjValorPro = GCDTMFECHANULA
        End If
    End Sub
    Private Function FblnAplicaDsctoPP() As Boolean
        Dim lblnAplica As Boolean
        lblnAplica = GobjParametros.ObjAnoActual.ObjAplicaDsctoPPBln.ObjValorPro AndAlso
                 ObjIdModoFacturacionByt.ObjValorPro = EnuModoFacturacionDef.EnuSistema
        If lblnAplica Then
            For Each lobjItemFactura As ClsItemFactura In ColItemsFactura
                lblnAplica = lobjItemFactura.ObjServicio.BlnEsCuotaAdministracion
                If lblnAplica Then Exit For
            Next
        End If
        Return lblnAplica
    End Function
    Private Sub SEstablezcaFechas()
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
            If ObjIdModoFacturacionByt.ObjValorPro = EnuModoFacturacionDef.EnuManual Then
                SEstablezcaFechasFM()
            Else
                SEstablezcaFechasFA()
            End If
        End If
    End Sub
    ''' <summary>
    ''' Establece las fechas para las facturas automáticas
    ''' </summary>
    Private Sub SEstablezcaFechasFA()
        Dim ldtmFechaGracia = GCDTMFECHANULA, ldtmFechaVence = GCDTMFECHANULA
        For Each lobjItemFact As ClsItemFactura In ColItemsFactura
            If ObjIdFormaPagoByt.ObjValorPro = EnuFormaPago.EnuContado Then
                ldtmFechaGracia = lobjItemFact.ObjFechaGraciaIFDtm.ObjValorPro
                ldtmFechaVence = lobjItemFact.ObjFechaVencimientoIFDtm.ObjValorPro
                Exit For
            Else
                If ldtmFechaGracia = GCDTMFECHANULA Then
                    ldtmFechaGracia = lobjItemFact.ObjFechaGraciaIFDtm.ObjValorPro
                ElseIf ldtmFechaGracia <> lobjItemFact.ObjFechaGraciaIFDtm.ObjValorPro Then
                    If lobjItemFact.ObjFechaGraciaIFDtm.ObjValorPro < ldtmFechaGracia Then
                        ldtmFechaGracia = lobjItemFact.ObjFechaGraciaIFDtm.ObjValorPro
                    End If
                End If
                If ldtmFechaVence = GCDTMFECHANULA Then
                    ldtmFechaVence = lobjItemFact.ObjFechaVencimientoIFDtm.ObjValorPro
                ElseIf ldtmFechaVence <> lobjItemFact.ObjFechaVencimientoIFDtm.ObjValorPro Then
                    If lobjItemFact.ObjFechaVencimientoIFDtm.ObjValorPro < ldtmFechaVence Then
                        ldtmFechaVence = lobjItemFact.ObjFechaVencimientoIFDtm.ObjValorPro
                    End If
                End If
            End If
        Next
        ObjFechaVencimientoDtm.ObjValorPro = ldtmFechaVence
        If ldtmFechaGracia < ObjFechaVencimientoDtm.ObjValorPro Then
            ldtmFechaGracia = ObjFechaVencimientoDtm.ObjValorPro
        End If
        ObjFechaGraciaDtm.ObjValorPro = ldtmFechaGracia
    End Sub
    ''' <summary>
    ''' Establece las fechas para las facturas manuales
    ''' </summary>
    Private Sub SEstablezcaFechasFM()
        Dim ldtmFechaVen As Date = ObjFechaVencimientoDtm.ObjValorPro
        Dim ldtmFechaGracia = GCDTMFECHANULA
        For Each lobjItemFact As ClsItemFactura In McolItemsFactura
            lobjItemFact.ObjFechaGraciaIFDtm.ObjValorPro =
                    lobjItemFact.ObjServicio.FdtmFechaGracias(ldtmFechaVen)
            If ldtmFechaGracia = GCDTMFECHANULA OrElse ldtmFechaGracia >
                    lobjItemFact.ObjFechaGraciaIFDtm.ObjValorPro Then
                ldtmFechaGracia = lobjItemFact.ObjFechaGraciaIFDtm.ObjValorPro
            End If
        Next
        ObjFechaGraciaDtm.ObjValorPro = ldtmFechaGracia
    End Sub
    Friend Sub SModifiqueADefinitiva()
        Dim lstrPrefijo = GobjParametros.FstrPrefijoDoc(EnuTipoDocOri.EnuFactura)
        HblnEsModificable = True
        McolItemsFactura = ColItemsFactura
        EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
        ObjEsPreFacturaBln.ObjValorPro = False
        ObjPrefijo_FactStr.ObjValorPro = lstrPrefijo
        If GobjParametros.BlnEFacAutorizado Then
            ObjPieFacturaUno_FactStr.ObjValorPro =
                    GobjParametros.ObjPieFacturaUnoStr.ObjValorPro
            ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuNoReg
        End If
        For Each lobjItemFactura As ClsItemFactura In ColItemsFactura
            lobjItemFactura.SModifiqueADefinitiva(lstrPrefijo)
        Next
        SActualice(True)
    End Sub
    Private Function FblnPuedeReversarItemProFac() As Boolean
        Dim lblnEsProgramacionFact = False
        For Each lobjItemFac As ClsItemFactura In ColItemsFactura
            lblnEsProgramacionFact = lobjItemFac.ObjServicio.ObjEsFactProgramableBln.ObjValorPro
            If lblnEsProgramacionFact Then Exit For
        Next
        Dim lstrPeriodoFac = ClsPanorama.FstrPeriodo(ObjFechaFacturaDtm.ObjValorPro)
        Dim lstrPeriodoActual = GobjParametros.ObjAnoActual.StrIdPeriodoActual
        Dim lblnFactEsPeridoAct = (lstrPeriodoFac = lstrPeriodoActual)
        Dim lblnPuede = lblnEsProgramacionFact AndAlso lblnFactEsPeridoAct
        If Not lblnPuede Then
            lblnPuede = Not lblnEsProgramacionFact
        End If
        Return lblnPuede
    End Function
    Friend Function FblnPredioAgrupadorEsValido(astrIdPredioAgrupador As String) As Boolean
        Dim lblnEsValido = True
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
            If Not IsNothing(McolItemsFactura) AndAlso McolItemsFactura.Count > 0 Then
                lblnEsValido = astrIdPredioAgrupador = ObjIdPredioAgrupador_FacStr.ObjValorPro
            End If
        End If
        Return lblnEsValido
    End Function
    Friend Function FbytQRFact() As Byte()
        Dim lentColorFondoQR As Integer = Color.FromArgb(255, 255, 255, 255).ToArgb()
        Dim lentColorQR As Integer = Color.FromArgb(255, 0, 0, 0).ToArgb()
        Dim lqreQRFact As New QRCodeEncoder With {
            .QRCodeEncodeMode = QRCodeEncoder.ENCODE_MODE.BYTE,
            .QRCodeScale = Int32.Parse(4),
            .QRCodeErrorCorrect = QRCodeEncoder.ERROR_CORRECTION.H,
            .QRCodeVersion = 0,
            .QRCodeBackgroundColor = System.Drawing.Color.FromArgb(lentColorFondoQR),
            .QRCodeForegroundColor = System.Drawing.Color.FromArgb(lentColorQR)
       }
        Dim lstrNit As String = GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual.
                ObjIdTerceroCentroUtilDbl.ToString
        Dim lbytQR As Byte() = Array.Empty(Of Byte)()
        Dim ldtmFecFac As Date = ObjFechaCreacionDtm.ObjValorPro
        Dim lstrFecFac As String = Year(ldtmFecFac).ToString &
                Format(Month(ldtmFecFac), "00") & Format(Day(ldtmFecFac), "00") &
                Format(Hour(ldtmFecFac), "00") & Format(Minute(ldtmFecFac), "00") &
                Format(Second(ldtmFecFac), "00")
        Dim lstrFac As String = "NumFac:" & StrNumeroFactura & vbCrLf &
                "FecFac:" & lstrFecFac & vbCrLf &
                "NitFac:" & lstrNit & vbCrLf &
                "DocAdq:" & ObjIdCliente_FactDbl.ToString & vbCrLf &
                "ValFac:" & Format(FdecValorServicios, "#0.00") & vbCrLf &
                "ValIva:" & Format(FdecIvaServicios, "#0.00") & vbCrLf &
                "ValOtroIm:" & "0.00" & vbCrLf &
                "ValFacIm:" & Format(ObjValor_FactDec.ObjValorPro, "#0.00") & vbCrLf
        If GobjParametros.ObjIdProveedorEFacEnt.ObjValorPro > EnuProveedorEFac.None Then
            lstrFac &= "CUFE:" & ObjCUFEStr.ObjValorPro
        End If
        Try
            Dim lbmiQRFac As New BitmapImage
            lbmiQRFac.BeginInit()
            Dim lbtmQR As Bitmap = lqreQRFact.Encode(lstrFac, System.Text.Encoding.UTF8)
            lbytQR = FBytQR(lbtmQR)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly)
        End Try
        Return lbytQR
    End Function
    Private Shared Function FBytQR(abmQR As Bitmap) As Byte()
        Dim lbytQR As Byte() = Array.Empty(Of Byte)()
        If abmQR IsNot Nothing Then
            Using lmstImagenBinaria As New MemoryStream
                abmQR.Save(lmstImagenBinaria, Imaging.ImageFormat.Jpeg)
                lbytQR = lmstImagenBinaria.GetBuffer
            End Using
        End If
        Return lbytQR
    End Function
    Friend Function FdecValorPagado() As Decimal
        Dim ldecVlrPagado = 0D
        For Each lobjItemFac As ClsItemFactura In ColItemsFactura
            ldecVlrPagado += lobjItemFac.FdecPagoTotalAplicado
        Next
        Return ldecVlrPagado
    End Function
    Friend Function FarlCorreosFac() As ArrayList
        Dim larlListaCorreos As New ArrayList
        Dim lstrCorreoCli As String, lstrCorreoPredio As String
        If ObjClienteFactura.ObjRecibeDocsPorEmailBln.ObjValorPro Then
            lstrCorreoCli = ObjClienteFactura.ObjEmailStr.ToString
            If Not String.IsNullOrEmpty(lstrCorreoCli) Then
                larlListaCorreos.Add(lstrCorreoCli)
            End If
        End If
        If Not IsNothing(ObjPredioAgrFactura) Then
            lstrCorreoPredio = ObjPredioAgrFactura.ObjEmailAdiStr.ToString
            If Not String.IsNullOrEmpty(lstrCorreoPredio) Then
                larlListaCorreos.Add(lstrCorreoPredio)
            End If
        End If
        Return larlListaCorreos
    End Function
    Friend Function FstrAliasCon() As String
        Dim lstrAliasCon = String.Empty
        If ObjPredioAgrFactura IsNot Nothing Then
            lstrAliasCon = ObjPredioAgrFactura.ObjAliasContStr.ToString
        End If
        If String.IsNullOrEmpty(lstrAliasCon) Then
            lstrAliasCon = ObjIdCliente_FactDbl.ToString
        End If
        Return lstrAliasCon
    End Function
    Friend Function FblnSerIdAplicado() As Boolean
        Dim lblnSerIdAplicado = False
        Dim lobjSerId As ClsServicio = GobjParametros.FobjServicioId
        Dim lshrIdSerId As Short = lobjSerId.ObjIdServicioShr.ObjValorPro
        For Each lobjItemFac As ClsItemFactura In ColItemsFactura
            lblnSerIdAplicado = lobjItemFac.ObjIdAno_ServicioItemFactShr.ObjValorPro = 0 AndAlso
                    lobjItemFac.ObjIdServicio_ItemFactShr.ObjValorPro = lshrIdSerId
            If lblnSerIdAplicado Then
                Exit For
            End If
        Next
        Return lblnSerIdAplicado
    End Function
    Friend Function FstrEmailCli() As String
        Return ObjClienteFactura.ObjEmailStr.ToString
    End Function
    ''' <summary>
    ''' Verifica que una factura nueva es integra
    ''' </summary>
    ''' <returns></returns>
    Friend Function FblnFacturaIntegra() As Boolean
        Dim ldecValorFacItems = 0D, ldecDebitosItems = 0D, i = 0, lblnEsIntegra = False
        If ObjValor_FactDec.ObjValorPro = ObjDebitos_FactDec.ObjValorPro Then
            For Each lobjItemFac As ClsItemFactura In ColItemsFactura
                i += 1
                lblnEsIntegra = lobjItemFac.ObjIdItemFacturaShr.ObjValorPro = i
                If lblnEsIntegra Then
                    ldecValorFacItems += lobjItemFac.ObjValor_ItemFactDec.ObjValorPro
                    ldecDebitosItems += lobjItemFac.ObjValor_ItemFactDec.ObjValorPro
                Else
                    Exit For
                End If
            Next
            If lblnEsIntegra Then
                lblnEsIntegra = ldecValorFacItems = ObjValor_FactDec.ObjValorPro AndAlso
                    ldecDebitosItems = ObjValor_FactDec.ObjValorPro
            End If
        End If
        Return lblnEsIntegra
    End Function
#End Region
#Region "Valores del servicio"
    ''' <summary>
    ''' Devuelve el valor total de los servicios sin tener en cuenta el IVA
    ''' </summary>
    ''' <returns></returns>
    Friend Function FdecValorServicios() As Decimal
        Dim ldecVlr = 0D
        For Each lobjItemFac As ClsItemFactura In ColItemsFactura
            ldecVlr += lobjItemFac.FdecValorServicio
        Next
        Return ldecVlr
    End Function
    Friend Function FdecBaseIvaCapital() As Decimal
        Dim ldecValoBaserIva = 0D
        For Each lobjItemFact As ClsItemFactura In ColItemsFactura
            ldecValoBaserIva += lobjItemFact.FdecBaseIvaServicio()
        Next
        Return ldecValoBaserIva
    End Function
    Friend Function FdecIvaServicios() As Decimal
        Dim ldecValorIva = 0D
        For Each lobjItemFact As ClsItemFactura In ColItemsFactura
            ldecValorIva += lobjItemFact.FdecIvaServicio
        Next
        Return ldecValorIva
    End Function
    Friend Function FdecBaseIvaInt() As Decimal
        Dim ldecBaseIvaInt = 0D
        For Each lobjItemFac As ClsItemFactura In ColItemsFactura
            ldecBaseIvaInt += lobjItemFac.FdecBaseIvaInt
        Next
        Return ldecBaseIvaInt
    End Function
    Friend Function FdecIvaInt() As Decimal
        Dim ldecIvaInt = 00D
        For Each lobjItemFact As ClsItemFactura In ColItemsFactura
            ldecIvaInt += lobjItemFact.FdecIvaInt
        Next
        Return ldecIvaInt
    End Function
    Friend Function FdecIvaTotal() As Decimal
        Return FdecIvaServicios() + FdecIvaInt()
    End Function
    Friend Function FdecBaseIvaTotal() As Decimal
        Return FdecBaseIvaCapital() + FdecBaseIvaInt()
    End Function
#End Region
#Region "Deuda"
    ''' <summary>
    ''' Devuelve el valor debido (debitos menos creditos).
    ''' </summary>
    Friend ReadOnly Property DecDeuda As Decimal
        Get
            Dim ldecDeudaFact As Decimal = ObjDebitos_FactDec.ObjValorPro -
                    ObjCreditos_FactDec.ObjValorPro
            Return ldecDeudaFact
        End Get
    End Property
    ''' <summary>
    ''' Devuelve el valor de la deuda del servicio pasado en el argumento
    ''' </summary>
    ''' <param name="astrServicio">Donde A = Todos, 0 = Administración, # = Servicio Permanente
    ''' identificado con el enero pasado en # </param>
    ''' <returns></returns>
    Friend ReadOnly Property DecDeudaServicio(astrServicio As String) As Decimal
        Get
            Dim ldecDeuda = 0D
            If astrServicio = "A" Then
                ldecDeuda = DecDeuda
            Else
                For Each lobjItemFac As ClsItemFactura In ColItemsFactura
                    If lobjItemFac.StrServicio = astrServicio Then
                        ldecDeuda += lobjItemFac.DecDeuda
                    End If
                Next
            End If
            Return ldecDeuda
        End Get
    End Property
    ''' <summary>
    ''' Devuelve el valor de la deuda de capital a partir de las novedades de los items de factura
    ''' incluyendo el IVA
    ''' </summary>
    Friend Function FdecDeudaCapital() As Decimal
        Dim ldecDeudaCapital = 0
        For Each lobjItemFac As ClsItemFactura In ColItemsFactura
            ldecDeudaCapital += lobjItemFac.FdecDeudaServicioTotal
        Next
        Return ldecDeudaCapital
    End Function
    Friend Function FdecDeudaCapitalAntesIva() As Decimal
        Dim ldecDeudaCapitalAntesIva = 0
        For Each lobjItemFac As ClsItemFactura In ColItemsFactura
            ldecDeudaCapitalAntesIva += lobjItemFac.FdecDeudaServicioTotal -
                    lobjItemFac.FdecDeudaIva
        Next
        Return ldecDeudaCapitalAntesIva
    End Function
    Friend ReadOnly Property DecDeudaCapitalSer(astrServicios As String()) As Decimal
        Get
            Dim ldecDeuda = 0D
            For Each lobjItemFac As ClsItemFactura In ColItemsFactura
                If astrServicios.Contains("A") Then
                    ldecDeuda += lobjItemFac.FdecDeudaServicioTotal
                Else
                    If lobjItemFac.ObjIdAno_ServicioItemFactShr.ObjValorPro > 0 Then
                        If astrServicios.Contains("0") Then
                            ldecDeuda += lobjItemFac.FdecDeudaServicioTotal
                        End If
                    Else
                        If astrServicios.Contains(lobjItemFac.ObjIdServicio_ItemFactShr.ToString()) Then
                            ldecDeuda += lobjItemFac.FdecDeudaServicioTotal
                        End If
                    End If
                End If
            Next
            Return ldecDeuda
        End Get
    End Property
    Friend ReadOnly Property DecDeudaIntMoraSer(astrServicios As String()) As Decimal
        Get
            Dim ldecDeuda = 0D
            For Each lobjItemFac As ClsItemFactura In ColItemsFactura
                If astrServicios.Contains("A") Then
                    ldecDeuda += lobjItemFac.FdecDeudaIntMora
                Else
                    If lobjItemFac.ObjIdAno_ServicioItemFactShr.ObjValorPro > 0 Then
                        If astrServicios.Contains("0") Then
                            ldecDeuda += lobjItemFac.FdecDeudaIntMora
                        End If
                    Else
                        If astrServicios.Contains(lobjItemFac.ObjIdServicio_ItemFactShr.ToString()) Then
                            ldecDeuda += lobjItemFac.FdecDeudaIntMora
                        End If
                    End If
                End If
            Next
            Return ldecDeuda
        End Get
    End Property
    ''' <summary>
    ''' Devuelve el valor de la deuda por mora a partir de las novedades
    ''' </summary>
    Friend Function FdecDeudaIntMora() As Decimal
        Dim ldecDeudaMora As Decimal
        For Each lobjItemFac As ClsItemFactura In ColItemsFactura
            ldecDeudaMora += lobjItemFac.FdecDeudaIntMora
        Next
        Return ldecDeudaMora
    End Function
    Friend Function FdecDeudaIntMoraAntesIva() As Decimal
        Dim ldecDeudaMoraAntesIva = 0D
        For Each lobjItemFac As ClsItemFactura In ColItemsFactura
            ' AVV Revisar
            'ldecDeudaMoraAntesIva += lobjItemFac.FdecDeudaIntMora - lobjItemFac.FdecDeudaIvaInt
        Next
        Return ldecDeudaMoraAntesIva
    End Function
    Friend Function FdecDeudaIva() As Decimal
        Dim ldecDeudaIva As Decimal
        For Each lobjItemFac As ClsItemFactura In ColItemsFactura
            ldecDeudaIva += lobjItemFac.FdecDeudaIva
        Next
        Return ldecDeudaIva
    End Function
#End Region
#Region "eFac"
    ''' <summary>
    ''' Indica si la factura es una factura electrónica
    ''' </summary>
    Friend ReadOnly Property BlnEsFacEle As Boolean
        Get
            Return GobjParametros.BlnEFacAutorizado AndAlso ObjIdEstadoEDocEnt.ObjValorPro <
                    EnuEstadoEDoc.EnuNoEDoc
        End Get
    End Property
    ''' <summary>
    ''' Indica si es una factura electrónica y si esta registrada 
    ''' ''' </summary>
    Friend ReadOnly Property BlnEstaRegEFac As Boolean
        Get
            Return BlnEsFacEle AndAlso ObjIdEstadoEDocEnt.ObjValorPro >= EnuEstadoEDoc.EnuRegi
        End Get
    End Property
    ''' <summary>
    ''' Indica si es una factura electrónica y esta por ser registrada 
    ''' ''' </summary>
    Friend ReadOnly Property BlnEstaPorRegEFac As Boolean
        Get
            Return BlnEsFacEle AndAlso ObjIdEstadoEDocEnt.ObjValorPro < EnuEstadoEDoc.EnuRegi
        End Get
    End Property
    Friend Function FenuVerFacEFac() As EnuVerEFac
        Dim lenuVerEFac As EnuVerEFac
        If BlnEstaRegEFac Then
            If ObjFechaFacturaDtm.ObjValorPro < DateSerial(2019, 12, 20) Then
                lenuVerEFac = EnuVerEFac.EnuV1
            Else
                lenuVerEFac = EnuVerEFac.EnuV2
            End If
        Else
            If BlnEsFacEle Then
                lenuVerEFac = EnuVerEFac.EnuV2
            Else
                lenuVerEFac = EnuVerEFac.EnuNinguna
            End If
        End If
        Return lenuVerEFac
    End Function
    Friend Function FblnInsertarEFac() As Boolean
        Dim lblnReg = False
        If GobjParametros.BlnEFacAutorizado Then
            Dim lobjEstadoFacEDoc As EnuEstadoEDoc = ObjIdEstadoEDocEnt.ObjValorPro
            lblnReg = (lobjEstadoFacEDoc = EnuEstadoEDoc.EnuNoReg)
        End If
        Return lblnReg
    End Function
    Friend Function FblnActualizarEstEFac() As Boolean
        Dim lblnActu = False
        If GobjParametros.BlnEFacAutorizado() Then
            lblnActu = ObjIdEstadoEDocEnt.ObjValorPro < EnuEstadoEDoc.EnuEnviada AndAlso
                    ObjIdEstadoEDocEnt.ObjValorPro >= EnuEstadoEDoc.EnuEnProceso
        End If
        Return lblnActu
    End Function
    Friend Function FblnEnviarEFac() As Boolean
        Dim lblnEnv = False
        If GobjParametros.BlnEFacAutorizado Then
            lblnEnv = ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuRegi
        End If
        Return lblnEnv
    End Function
    Friend Sub SHabiliteProcesarEFac()
        If GobjParametros.BlnEFacAutorizado Then
            If ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuInvalida OrElse
                    ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuErrorFtp Then
                If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
                    EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
                End If
                ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuNoReg
                SActualice(True)
            End If
        End If
    End Sub
    ''' <summary>
    ''' Prepara la factura para ser reenvaida al proveedor de facturación electrónica 
    ''' </summary>
    Friend Sub SPrepareParaReprocesarEfac()
        If GobjParametros.ObjAutorizaEFacBln.ObjValorPro AndAlso ObjIdEstadoEDocEnt.ObjValorPro <>
                EnuEstadoEDoc.EnuNoReg AndAlso ObjIdEstadoEDocEnt.ObjValorPro <>
                EnuEstadoEDoc.EnuNoEDoc AndAlso ObjIdEstadoEDocEnt.ObjValorPro <>
                EnuEstadoEDoc.EnuInvalida Then
            If Not String.IsNullOrEmpty(ObjCUDocStr.ToString) Then
                EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
                ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuErrorFtp
                ObjCUFEStr.ObjValorPro = ""
                ObjCUDocStr.ObjValorPro = ""
                SActualice(True)
            End If
        End If
    End Sub
#End Region
#Region "Manejo Items Factura"
    Friend ReadOnly Property ColItemsFactura As Collection
        Get
            If IsNothing(McolItemsFactura) OrElse McolItemsFactura.Count = 0 Then
                McolItemsFactura = New Collection
                If ObjIdFacturaEnt.BlnEsValido AndAlso EnuEstadoActualizacion <> EnuEstadoObjetoDef.EnuCreando Then
                    SCargueDtbItems()
                    If Not IsNothing(MdtbItemsFact) AndAlso MdtbItemsFact.Rows.Count > 0 Then
                        Dim ldrwItemsFact() As DataRow = MdtbItemsFact.Select
                        For Each ldrwItemFact As DataRow In ldrwItemsFact
                            Dim lobjItemFact As New ClsItemFactura(Me, ldrwItemFact)
                            lobjItemFact.SLeaValores(True)
                            McolItemsFactura.Add(lobjItemFact,
                                    lobjItemFact.ObjIdItemFacturaShr.ToString)
                        Next
                    End If
                End If
            End If
            Return McolItemsFactura
        End Get
    End Property
    Friend Function FobjNuevoItemFactura() As ClsItemFactura
        Dim lobjItemFact As ClsItemFactura = Nothing
        Try
            SCargueDtbItems()
            Dim ldrwNewItemFact = MdtbItemsFact.NewRow
            lobjItemFact = New ClsItemFactura(Me, ldrwNewItemFact)
            With lobjItemFact
                .SCreeObj(Nothing)
                .ObjPrefijo_ItemFactStr.ObjValorPro = Me.ObjPrefijo_FactStr.ObjValorPro
            End With
        Catch ex As ErrorInesperadoPanDatException
            Throw
        Catch ex As ErrorInesperadoPanLException
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
        Return lobjItemFact
    End Function
    Friend Sub SAdicioneNuevoItem(aobjNuevoItemFactura As ClsItemFactura)
        Try
            If IsNothing(McolItemsFactura) Then
                McolItemsFactura = ColItemsFactura
            End If
            Dim lshrIdItemFac As Short = McolItemsFactura.Count + 1
            With aobjNuevoItemFactura
                .ObjIdItemFacturaShr.ObjValorPro = lshrIdItemFac
                .ObjDebitos_ItemFactDec.ObjValorPro = .ObjValor_ItemFactDec.ObjValorPro
                .ObjEsExcluidoIva_ItemFactBln.ObjValorPro = .ObjServicio.ObjEsExcluidoIvaBln.ObjValorPro
                .ObjTarifaIva_ItemFactDbl.ObjValorPro = .ObjServicio.ObjTarifaIvaDbl.ObjValorPro
            End With
            McolItemsFactura.Add(aobjNuevoItemFactura)
            SEstablezcaPiePagina()
            SCalculeValorFactura()
        Catch ex As ErrorInesperadoPanDatException
            Throw
        Catch ex As ErrorInesperadoPanLException
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
    End Sub
    Friend Sub SAdicioneItemSerId(aobjNuevoItemFactura As ClsItemFactura)
        Try
            If IsNothing(McolItemsFactura) Then
                McolItemsFactura = ColItemsFactura
            End If
            With aobjNuevoItemFactura
                .ObjDebitos_ItemFactDec.ObjValorPro = .ObjValor_ItemFactDec.ObjValorPro
                .ObjEsExcluidoIva_ItemFactBln.ObjValorPro = .ObjServicio.ObjEsExcluidoIvaBln.ObjValorPro
                .ObjTarifaIva_ItemFactDbl.ObjValorPro = .ObjServicio.ObjTarifaIvaDbl.ObjValorPro
                .ObjIdItemFacturaShr.ObjValorPro = McolItemsFactura.Count + 1
                .ObjPrefijo_ItemFactStr.ObjValorPro = ObjPrefijo_FactStr.ObjValorPro
                .ObjIdFactura_ItemFactEnt.ObjValorPro = ObjIdFacturaEnt.ObjValorPro
                .ObjEsPrefactura_ItemFactBln.ObjValorPro = ObjEsPreFacturaBln.ObjValorPro
            End With
            McolItemsFactura.Add(aobjNuevoItemFactura)
            SCalculeValorFactura()
        Catch ex As ErrorInesperadoPanDatException
            Throw
        Catch ex As ErrorInesperadoPanLException
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
    End Sub
    ''' <summary>
    ''' Elimina un item de la factura.
    ''' </summary>
    ''' <param name="ashrIdItemFra"></param>
    ''' <remarks>Solo se usa para eliminar un item de la factura manual en creación</remarks>
    Friend Sub SElimineItem(ashrIdItemFra As Short)
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
            If McolItemsFactura.Count > 0 AndAlso McolItemsFactura.Count >= ashrIdItemFra Then
                McolItemsFactura.Remove(ashrIdItemFra)
                SRenumereItems()
                SCalculeValorFactura()
            Else
                Throw New ErrorInesperadoPanLException("Objeto item factura no encontrado!")
            End If
        End If
    End Sub
    Private Sub SRenumereItems()
        Dim i = 0
        For Each lobjItemFac As ClsItemFactura In ColItemsFactura
            i += 1
            lobjItemFac.ObjIdItemFacturaShr.ObjValorPro = i
            SRenumereNovedades(lobjItemFac)
        Next
    End Sub
    Private Sub SRenumereNovedades(aobjItemFac As ClsItemFactura)
        Dim i = 0
        For Each lobjNov As ClsNovedad In aobjItemFac.ColNovedades
            i += 1
            lobjNov.ObjIdItemFact_NovShr.ObjValorPro = aobjItemFac.ObjIdItemFacturaShr.ObjValorPro
            lobjNov.ObjIdNovedadShr.ObjValorPro = i
        Next
    End Sub
    Private Sub SCalculeValorFactura()
        Dim ldecvalorFac = 0D
        If Not IsNothing(McolItemsFactura) AndAlso McolItemsFactura.Count > 0 Then
            For Each lobjItemFac As ClsItemFactura In McolItemsFactura
                ldecvalorFac += lobjItemFac.ObjValor_ItemFactDec.ObjValorPro
            Next
        End If
        ObjValor_FactDec.ObjValorPro = ldecvalorFac
        ObjDebitos_FactDec.ObjValorPro = ldecvalorFac
    End Sub
    Private Sub SEstablezcaPiePagina()
        If Not IsNothing(McolItemsFactura) AndAlso McolItemsFactura.Count > 0 Then
            ObjPieFacturaUno_FactStr.ObjValorPro = GobjParametros.ObjPieFacturaUnoStr.ToString
            ObjPieFacturaDos_FactStr.ObjValorPro = GobjParametros.ObjPieFacturaDosStr.ToString
        End If
    End Sub
    ''' <summary>
    ''' Devuelve la datatable con los items de la factura actual
    ''' </summary>
    ''' <returns>DataTable</returns>
    Friend ReadOnly Property DtbItemsFact As DataTable
        Get
            SCargueDtbItems()
            SComplementeTablaItems()
            Return MdtbItemsFact
        End Get
    End Property
    Private Sub SCargueDtbItems()
        If IsNothing(MdtbItemsFact) Then
            Dim lstrIdFactura = "0"
            Dim lstrPrefijo = String.Empty
            If EnuEstadoActualizacion <> EnuEstadoObjetoDef.EnuCreando Then
                lstrIdFactura = ObjIdFacturaEnt.ToString
                If String.IsNullOrEmpty(lstrIdFactura) Then lstrIdFactura = "0"
                lstrPrefijo = ObjPrefijo_FactStr.ToString
            End If
            Dim lstrIndice = {{StrCampoCarpeta, "ASC"},
                              {StrCampoCentroUtil, "ASC"},
                              {ClsPrefijo_FactStr.SstrNombreCampoBd, "ASC"},
                              {ClsIdFactura_ItemFactEnt.SstrNombreCampoBd, "ASC"},
                              {ClsIdItemFacturaShr.SstrNombreCampoBd, "ASC"}}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsPrefijo_FactStr.SstrNombreCampoBd & " = '" & lstrPrefijo &
                    "' AND " & ClsIdFactura_ItemFactEnt.SstrNombreCampoBd & " = " & lstrIdFactura
            Dim lstrCamposSelect() = {"*", "0 as VlrIva"}
            MdtbItemsFact = ClsPanorama.FdtbDataTable(ClsItemFactura.SstrNombreTabla, lstrCamposSelect, lstrIndice,
                    lstrFiltro)
        End If
    End Sub
    Private Sub SComplementeTablaItems()
        If MstrNroFactItem <> StrNumeroFactura Then
            MstrNroFactItem = StrNumeroFactura
            If IsNothing(McolItemsFactura) Then
                McolItemsFactura = ColItemsFactura
            End If
            Dim lobjcenUtil As ClsCentroUtilidad = GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual
            Dim ldrwItemsFactura As DataRow() = MdtbItemsFact.Select()
            Dim lobjItemFac As ClsItemFactura = Nothing
            Dim ldecVlrIvaItem As Decimal
            For Each ldrwItFa As DataRow In ldrwItemsFactura
                Dim lshrIdItemFac As Short = ClsPanorama.FobjValorCampo(ldrwItFa(ClsIdItemFacturaShr.SstrNombreCampoBd),
                            EnuTipoValor.EnuShort)
                If McolItemsFactura.Contains(lshrIdItemFac.ToString) Then
                    lobjItemFac = McolItemsFactura(lshrIdItemFac.ToString)
                End If
                ldecVlrIvaItem = lobjItemFac.FdecIvaServicio
                ldrwItFa("VlrIva") = ldecVlrIvaItem
            Next
        End If
    End Sub
    ''' <summary>
    ''' Devuelve una DataTable de los items de la factura sin datos, solo con la estructura
    ''' </summary>
    ''' <returns>DataTable</returns>
    ''' <remarks>Se utiliza solo para llenar el datagrid de la forma winFacturaManual</remarks>
    Friend Function FdtbItems() As DataTable
        GobjPanDat.SControleProcesoObj(True)
        Dim ldtbItems As DataTable
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsPrefijo_FactStr.SstrNombreCampoBd & " = '" & ObjPrefijo_FactStr.ObjValorPro &
                "' AND " & ClsIdFacturaEnt.SstrNombreCampoBd & " = " & "0"
        Dim lstrCamposSelect() = {"*"}
        ldtbItems = ClsPanorama.FdtbDataTable(ClsItemFactura.SstrNombreTabla, lstrCamposSelect, {{"", ""}},
                lstrFiltro)
        GobjPanDat.SControleProcesoObj(False)
        Return ldtbItems
    End Function
    Friend Function FdecDeudaPredio(astrIdpredio As String)
        Dim ldecDeuda = 0D
        If IsNothing(McolItemsFactura) Then
            McolItemsFactura = ColItemsFactura
        End If
        For Each lobjItemFac As ClsItemFactura In McolItemsFactura
            If lobjItemFac.ObjIdPredio_ItemFactStr.ObjValorPro = astrIdpredio Then
                ldecDeuda += lobjItemFac.DecDeuda
            End If
        Next
        Return ldecDeuda
    End Function
    ''' <summary>
    ''' Indica si el predio "astrIdPredio" tiene deuda en esta factura
    ''' </summary>
    ''' <param name="astrIdPredio">Id del Predio al cual se le ha facturado en esta factura</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function FblnPredioConDeuda(astrIdPredio As String) As Boolean
        Dim lblnFPredioConDeuda = False
        If IsNothing(McolItemsFactura) Then
            McolItemsFactura = ColItemsFactura
        End If
        For Each lobjItemFac As ClsItemFactura In McolItemsFactura
            If lobjItemFac.ObjIdPredio_ItemFactStr.ObjValorPro = astrIdPredio Then
                If lobjItemFac.DecDeuda > 0 Then
                    lblnFPredioConDeuda = True
                    Exit For
                End If
            End If
        Next
        Return lblnFPredioConDeuda
    End Function
    Friend Function FstrIdTerceroCtaCr(ashrIdItemFac As Short) As String
        Dim lobjItemFac As ClsItemFactura = ColItemsFactura(ashrIdItemFac.ToString)
        Dim lstrIdTer As String = lobjItemFac.ObjServicio.ObjIdTerceroCtaCrDbl.ToString
        If lstrIdTer = "0" Then
            lstrIdTer = String.Empty
        End If
        Return lstrIdTer
    End Function
    ''' <summary>
    ''' Indica si la factura tiene algun item del servicios permanente pasado en el argumento
    ''' </summary>
    ''' <returns></returns>
    Friend Function FblnFacturoServicio(ashrIdAno As Short, ashrIdServicio As Short)
        Dim lblnFactSer = False
        For Each lobjItemFac As ClsItemFactura In ColItemsFactura
            If lobjItemFac.ObjIdAno_ServicioItemFactShr.ObjValorPro = ashrIdAno Then
                lblnFactSer = lobjItemFac.ObjIdServicio_ItemFactShr.ObjValorPro = ashrIdServicio
            End If
        Next
        Return lblnFactSer
    End Function
#End Region
#Region "Iva y Retenciones"
    Friend Function FstrIvasFactura() As String()
        Dim i = -1
        Dim ldblTasasIva() As Double = Array.Empty(Of Double)(), ldblIvaItem As Double
        For Each lobjItemFac As ClsItemFactura In ColItemsFactura
            If lobjItemFac.FdecIvaServicio > 0 Then
                ldblIvaItem = lobjItemFac.ObjTarifaIva_ItemFactDbl.ObjValorPro
                If Not ldblTasasIva.Contains(ldblIvaItem) Then
                    i += 1
                    ReDim Preserve ldblTasasIva(i)
                    ldblTasasIva(i) = ldblIvaItem
                End If
            End If
        Next
        Dim lstrIvasFact(ldblTasasIva.Count - 1) As String
        i = -1
        For Each ldblTasaIva As Double In ldblTasasIva
            i += 1
            lstrIvasFact(i) = FstrValoresIva(ldblTasaIva)
        Next
        Return lstrIvasFact
    End Function
    Private Function FstrValoresIva(adblTasaIva As Double) As String
        Dim ldecTotaBase = 0D, ldecTotaVlrIva = 0D
        For Each lobjItemFac As ClsItemFactura In McolItemsFactura
            If lobjItemFac.ObjTarifaIva_ItemFactDbl.ObjValorPro = adblTasaIva Then
                ldecTotaBase += lobjItemFac.FdecBaseIvaServicio
                ldecTotaVlrIva += lobjItemFac.FdecIvaServicio
            End If
        Next
        Dim lstrValoresIva = Format(adblTasaIva * 100, "#0.00") & "&" & Format(ldecTotaBase, "#0.00") & "&" &
                Format(ldecTotaVlrIva, "#0.00")
        Return lstrValoresIva
    End Function
#End Region
#Region "Novedades Factura"
    Friend ReadOnly Property DtbNovedadesFact As DataTable
        Get
            SCargueDtbNovedadesFact()
            SComplementeTablaNov()
            Return MdtbNovedadesFact
        End Get
    End Property
    Private Sub SCargueDtbNovedadesFact()
        If IsNothing(MdtbNovedadesFact) Then
            Dim lstrIdFactura = ObjIdFacturaEnt.ToString
            If String.IsNullOrEmpty(lstrIdFactura) Then lstrIdFactura = "0"
            Dim lstrTabla = ClsNovedad.SstrNombreTabla
            Dim lstrCamposSelect = {"*", "'' AS ConceptoNovedad", "'' AS DocOrigen"}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsPrefijoFact_NovStr.SstrNombreCampoBd &
                    " = '" & ObjPrefijo_FactStr.ObjValorPro & "' AND " & ClsIdFactura_NovEnt.SstrNombreCampoBd &
                    " = " & lstrIdFactura & " AND " & ClsValor_NovDec.SstrNombreCampoBd & " <> 0"
            Dim lstrIndice = {{ClsIdFactura_NovEnt.SstrNombreCampoBd, "ASC"}, {ClsIdNovedadShr.SstrNombreCampoBd, "ASC"}}
            MdtbNovedadesFact = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamposSelect, lstrIndice, lstrFiltro)
        End If
    End Sub
    Private Sub SComplementeTablaNov()
        If MstrNroFactura <> StrNumeroFactura Then
            MstrNroFactura = StrNumeroFactura
            Dim ldrwNovedades = MdtbNovedadesFact.Select
            Dim lstrConceptoNovedad As String
            For Each ldrwNovedad As DataRow In ldrwNovedades
                Dim lenuTipoNovedad As EnuTipoNov =
                        ClsPanorama.FobjValorCampo(ldrwNovedad(ClsIdTipoNovedadByt.SstrNombreCampoBd),
                        EnuTipoValor.EnuByte)
                Dim lshrIdItemFact_Nov As Short =
                        ClsPanorama.FobjValorCampo(ldrwNovedad(ClsIdItemFacturaShr.SstrNombreCampoBd),
                        EnuTipoValor.EnuShort)
                Dim lenuDocOrigen As EnuTipoDocOri =
                        ClsPanorama.FobjValorCampo(ldrwNovedad(ClsIdTipoDocOrigenByt.SstrNombreCampoBd),
                        EnuTipoValor.EnuInteger)
                Dim lstrDocOrigen = ClsOrionCop.FstrDocOrigenNovedad(lenuDocOrigen)
                lstrConceptoNovedad = FstrConceptoNovedad(lenuTipoNovedad, lshrIdItemFact_Nov)
                ldrwNovedad("DocOrigen") = lstrDocOrigen
                ldrwNovedad("ConceptoNovedad") = lstrConceptoNovedad
            Next
        End If
    End Sub
    Friend Function FshrIdUltimaNov() As Short
        Dim lshrIdUltNov = 0S
        Dim lstrIdFactura = ObjIdFacturaEnt.ToString
        If String.IsNullOrEmpty(lstrIdFactura) Then lstrIdFactura = "0"
        Dim lstrTabla = ClsNovedad.SstrNombreTabla
        Dim lstrCamposSelect = {"*", "'' AS ConceptoNovedad", "'' AS DocOrigen"}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsPrefijoFact_NovStr.SstrNombreCampoBd &
                    " = '" & ObjPrefijo_FactStr.ObjValorPro & "' AND " & ClsIdFactura_NovEnt.SstrNombreCampoBd &
                    " = " & lstrIdFactura
        Dim lstrIndice = {{ClsIdFactura_NovEnt.SstrNombreCampoBd, "ASC"}, {ClsIdNovedadShr.SstrNombreCampoBd, "ASC"}}
        Dim ldtbNovedades = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamposSelect, lstrIndice, lstrFiltro)
        If ldtbNovedades.Rows.Count > 0 Then
            Dim ldrwUltNov As DataRow = ldtbNovedades.Rows(ldtbNovedades.Rows.Count - 1)
            lshrIdUltNov = ClsPanorama.FobjValorCampo(ldrwUltNov(ClsIdNovedadShr.SstrNombreCampoBd),
                    EnuTipoValor.EnuShort)
        End If
        Return lshrIdUltNov
    End Function
    Friend Function FstrConceptoNovedad(aenuTipoNovedad As EnuTipoNov,
                ashrIdItemFac As Short) As String 'Ok TipoNov
        Dim lstrConceptoNovedad = FstrConceptoNovedad(aenuTipoNovedad)
        If String.IsNullOrEmpty(lstrConceptoNovedad) Then
            Select Case aenuTipoNovedad
                Case EnuTipoNov.EnuDbCap
                    lstrConceptoNovedad = FstrNombreServicio(ashrIdItemFac)
                Case EnuTipoNov.EnuDbAntDev
                    lstrConceptoNovedad = "Anticipo reintegrado"
                Case EnuTipoNov.EnuDbAntApl
                    lstrConceptoNovedad = "Anticipo aplicado"
                Case EnuTipoNov.EnuRDbCap
                    lstrConceptoNovedad = FstrNombreServicio(ashrIdItemFac) & " Reversado"
                Case EnuTipoNov.EnuRDbIva
                    lstrConceptoNovedad = "Iva generado Reversado"
                Case EnuTipoNov.EnuCrIvaGas
                    lstrConceptoNovedad = "Iva llevado al Gasto"
                Case EnuTipoNov.EnuRDbInt
                    lstrConceptoNovedad = "Intereses de Mora Reversados"
                Case EnuTipoNov.EnuRCrPagoCap
                    lstrConceptoNovedad = "Abono a Capital Reversado"
                Case EnuTipoNov.EnuRCrPagoInt
                    lstrConceptoNovedad = "Abono a Intereses de Mora Reversado"
                Case EnuTipoNov.EnuRCrAnApCap
                    lstrConceptoNovedad = "Anticipo aplicado a Capital Reversado"
                Case EnuTipoNov.EnuRCrAnApInt
                    lstrConceptoNovedad = "Anticipo aplicado a Intereses de Mora Reversado"
                Case EnuTipoNov.EnuRCrDctoCap
                    lstrConceptoNovedad = "Devolución de Capital Reversado"
                Case EnuTipoNov.EnuRCrIvaGas
                    lstrConceptoNovedad = "Iva llevado al Gasto Reversado"
                Case EnuTipoNov.EnuRCrDctoInt
                    lstrConceptoNovedad = "Devolución de Intereses de Mora Reversado"
                Case EnuTipoNov.EnuRCrRetFte
                    lstrConceptoNovedad = "Retención en la Fuente Reversado"
                Case EnuTipoNov.EnuRCrRetIva
                    lstrConceptoNovedad = "Retencion del IVA Reversado"
                Case EnuTipoNov.EnuRCrRetIca
                    lstrConceptoNovedad = "Retención de Industria y Comercio Reversado"
                Case EnuTipoNov.EnuRCrRetCre
                    lstrConceptoNovedad = "Retencion del CREE Reversado"
                Case EnuTipoNov.EnuRCrAntRec
                    lstrConceptoNovedad = "Anticipo recibido Reversado"
                Case EnuTipoNov.EnuRDbAntDev
                    lstrConceptoNovedad = "Anticipo reintegrado Reversado"
                Case EnuTipoNov.EnuRDbAntApl
                    lstrConceptoNovedad = "Anticipo aplicado Reversado"
                Case EnuTipoNov.EnuDbIvaInt
                    lstrConceptoNovedad = "Iva Intereses de Mora"
                Case EnuTipoNov.EnuRDbIvaInt
                    lstrConceptoNovedad = "Iva Intereses de Mora Reversado"
            End Select
        End If
        Return lstrConceptoNovedad
    End Function
    Private Shared Function FstrConceptoNovedad(aenuTipoNovedad As EnuTipoNov) As String
        Dim lstrConceptoNovedad = String.Empty
        Select Case aenuTipoNovedad
            Case EnuTipoNov.EnuDbIva
                lstrConceptoNovedad = "Iva generado"
            Case EnuTipoNov.EnuDbIvaInt
                lstrConceptoNovedad = "Iva a Intereses de Mora"
            Case EnuTipoNov.EnuDbInt
                lstrConceptoNovedad = "Intereses de Mora causados"
            Case EnuTipoNov.EnuCrPagoCap
                lstrConceptoNovedad = "Abono a Capital"
            Case EnuTipoNov.EnuCrPagoInt
                lstrConceptoNovedad = "Abono a Intereses de Mora"
            Case EnuTipoNov.EnuCrAnApCap
                lstrConceptoNovedad = "Anticipo aplicado a Capital "
            Case EnuTipoNov.EnuCrAnApInt
                lstrConceptoNovedad = "Anticipo aplicado a Intereses de Mora "
            Case EnuTipoNov.EnuCrDctoCap
                lstrConceptoNovedad = "Devolución de Capital"
            Case EnuTipoNov.EnuCrIvaGas
                lstrConceptoNovedad = "IVA llevado al Gasto"
            Case EnuTipoNov.EnuCrDctoInt
                lstrConceptoNovedad = "Devolución de Intereses de Mora"
            Case EnuTipoNov.EnuCrRetFte
                lstrConceptoNovedad = "Retención en la Fuente"
            Case EnuTipoNov.EnuCrRetIva
                lstrConceptoNovedad = "Retencion del IVA"
            Case EnuTipoNov.EnuCrRetIca
                lstrConceptoNovedad = "Retención de Industria y Comercio"
            Case EnuTipoNov.EnuCrRetCre
                lstrConceptoNovedad = "Retencion del CREE"
            Case EnuTipoNov.EnuCrAntRec
                lstrConceptoNovedad = "Anticipo recibido"
        End Select
        Return lstrConceptoNovedad
    End Function
    Private Function FstrNombreServicio(ashrIdItemFactura As Short) As String
        Dim lobjItemFactura As ClsItemFactura = ColItemsFactura(ashrIdItemFactura.ToString)
        Dim lobjservicio As ClsServicio = lobjItemFactura.ObjServicio
        Return lobjservicio.ObjNombreServicioStr.ObjValorPro
    End Function
    ''' <summary>
    ''' Reversa en la factura el valor de la novedad pasada en el argumento aobjNovedad debido a la anulación 
    ''' de un recibo de caja que generó esta novedad o a la reversión de una nota contable generada apartir de 
    ''' un aticipo creado por un recibo de caja que se reverso.
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub SReverseNovedad(aobjNov As ClsNovedad,
                               aenuTipoDocOrigen As EnuTipoDocOri,
                               adtmFechaRev As Date,
                               astrPrefNotaRevCr As String,
                               aentIdNotaRevCr As Integer)
        Dim lshrIdItemFac As Short = aobjNov.ObjIdItemFact_NovShr.ObjValorPro
        Dim lobjItemFac As ClsItemFactura = ColItemsFactura(lshrIdItemFac.ToString)
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
            EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
        End If
        lobjItemFac.SReverseNovedad(aobjNov, adtmFechaRev, aenuTipoDocOrigen,
                astrPrefNotaRevCr, aentIdNotaRevCr, 0)
        ObjDebitos_FactDec.ObjValorPro += aobjNov.ObjValor_NovDec.ObjValorPro
        ObjFechaCancelacion_FactDtm.ObjValorPro = GCDTMFECHANULA
    End Sub
    ''' <summary>
    ''' Restaura la factura a los valores anteriores a la nota de reversion de recibo de caja 
    ''' por anulación de ésta.
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub SRestaureNovedadCr(aobjNovedad As ClsNovedad)
        Dim lshrIdItemFac As Short = aobjNovedad.ObjIdItemFact_NovShr.ObjValorPro
        Dim lobjItemFac As ClsItemFactura = ColItemsFactura(lshrIdItemFac.ToString)
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
            EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
        End If
        lobjItemFac.SReverseNotaReversionRC(aobjNovedad.ObjValor_NovDec.ObjValorPro,
                aobjNovedad.ObjFechaNovedadDtm.ObjValorPro)
        ObjCreditos_FactDec.ObjValorPro += aobjNovedad.ObjValor_NovDec.ObjValorPro
        If ObjDebitos_FactDec.ObjValorPro = ObjCreditos_FactDec.ObjValorPro Then
            ObjFechaCancelacion_FactDtm.ObjValorPro = aobjNovedad.ObjFechaNovedadDtm.ObjValorPro
        End If
    End Sub
#End Region
#Region "Calculo y causación intereses mora"
    ''' <summary>
    ''' Calcúla y devuelve los intereses de mora que se cobrarán desde la última fecha en que se 
    ''' causó mora  hasta la fecha pasada en el argumento 'adtmFechaCalculo'
    ''' </summary>
    ''' <param name="adtmFechaCalculo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend ReadOnly Property FdecIntMoraPorCausar(adtmFechaCalculo As Date)
        Get
            Dim ldecIntMoraFac = 0D, lblnCauseMora As Boolean
            If ObjPredioAgrFactura IsNot Nothing Then
                lblnCauseMora = ObjPredioAgrFactura.ObjIdEstadoDeuda_PredioByt.ObjValorPro <
                    EnuEstadoDeudaDef.EnuPerdida
            Else
                lblnCauseMora = ObjClienteFactura.ObjIdEstadoDeudaByt.ObjValorPro <
                        EnuEstadoDeudaDef.EnuPerdida
            End If
            If lblnCauseMora Then
                If FdecDeudaCapital() > 0 Then
                    For Each lobjItemFactura As ClsItemFactura In ColItemsFactura
                        If lobjItemFactura.FdecDeudaServicioTotal > 0 Then
                            ldecIntMoraFac += lobjItemFactura.FdecIntePorCausar(adtmFechaCalculo)
                        End If
                    Next
                End If
            End If
            Return ldecIntMoraFac
        End Get
    End Property
    Friend Property StcIntMora_DecBase(aentIndice As Integer) As Decimal
        Get
            Return MstcIntMora(aentIndice).DecBaseIntereses
        End Get
        Set(value As Decimal)
            MstcIntMora(aentIndice).DecBaseIntereses = value
        End Set
    End Property
    Friend Property StcIntMora_DecVlrMora(aentIndice As Integer) As Decimal
        Get
            Return MstcIntMora(aentIndice).DecVlrMora
        End Get
        Set(value As Decimal)
            MstcIntMora(aentIndice).DecVlrMora = value
        End Set
    End Property
    Friend Property StcIntMora_ShrIdItemFac(aentIndice As Integer) As Short
        Get
            Return MstcIntMora(aentIndice).ShrIdItemFactura
        End Get
        Set(value As Short)
            MstcIntMora(aentIndice).ShrIdItemFactura = value
        End Set
    End Property
    Friend Property StcIntMora_EntDiasMora(aentIndice As Integer) As Integer
        Get
            Return MstcIntMora(aentIndice).EntDiasMora
        End Get
        Set(value As Integer)
            MstcIntMora(aentIndice).EntDiasMora = value
        End Set
    End Property
    Friend Property StcIntMora_DtmFechaCauso(aentIndice As Integer) As Date
        Get
            Return MstcIntMora(aentIndice).DtmFechaCauso
        End Get
        Set(value As Date)
            MstcIntMora(aentIndice).DtmFechaCauso = value
        End Set
    End Property
    Friend Property StcIntMora_DblTasaIva(aentIndice As Integer) As Double
        Get
            Return MstcIntMora(aentIndice).DblTarifaIva
        End Get
        Set(value As Double)
            MstcIntMora(aentIndice).DblTarifaIva = value
        End Set
    End Property
    Friend ReadOnly Property StcIntMoraFactura(aentIndice As Integer) As StcIntMoraFactura
        Get
            Return MstcIntMora(aentIndice)
        End Get
    End Property
    ''' <summary>
    ''' Causa los intereses de mora a la factura a partir de la ultima vez que se hizo la causación.
    ''' </summary>
    ''' <param name="adtmFechaCausacion">Fecha a la cual se hace la causación.</param>
    ''' <returns>El valor de los intereses de mora causados.</returns>
    ''' <remarks></remarks>
    ' Causa mora a factura en procesos FM
    Friend Function SCauseMora(adtmFechaCausacion As Date) As Decimal
        ReDim MstcIntMora(ColItemsFactura.Count - 1)
        Dim ldecMoraFactura = 0D, lentDiasMora As Integer
        Dim ldecMoraItem As Decimal
        Dim i = -1
        For Each lobjItemFact As ClsItemFactura In ColItemsFactura
            ldecMoraItem = 0
            i += 1
            lobjItemFact.SCauseMora(adtmFechaCausacion, ldecMoraItem, lentDiasMora)
            ldecMoraFactura += ldecMoraItem
            StcIntMora_ShrIdItemFac(i) = lobjItemFact.ObjIdItemFacturaShr.ObjValorPro
            StcIntMora_DecVlrMora(i) = ldecMoraItem
            StcIntMora_DecBase(i) = lobjItemFact.FdecDeudaServicioTotal
            StcIntMora_EntDiasMora(i) = lentDiasMora
            StcIntMora_DtmFechaCauso(i) = lobjItemFact.ObjFechaCausoIntMora_Dtm.ObjValorPro
            StcIntMora_DblTasaIva(i) = lobjItemFact.ObjTarifaIva_ItemFactDbl.ObjValorPro
        Next
        If FblbModificoItem() Then
            SModifique()
            ObjDebitos_FactDec.ObjValorPro += ldecMoraFactura
        End If
        Return ldecMoraFactura
    End Function
    Friend Function FblbModificoItem() As Boolean
        Dim lblnModifico = False
        For Each lobjItemFac As ClsItemFactura In ColItemsFactura
            lblnModifico = lobjItemFac.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
            If lblnModifico Then Exit For
        Next
        Return lblnModifico
    End Function
    Friend Function FenuModoCausaMora() As EnuModoCausaMora
        Dim lenuModoCM As EnuModoCausaMora = EnuModoCausaMora.None
        For Each lobjItemFac As ClsItemFactura In ColItemsFactura
            If lobjItemFac.ObjServicio.BlnEsCuotaAdministracion Then
                lenuModoCM = lobjItemFac.ObjServicio.ObjModoCausaInteresesByt.ObjValorPro
            End If
        Next
        Return lenuModoCM
    End Function
#End Region
#Region "Aplica Creditos (Pagos, Descuentos, Retenciones y Anticipos)"
    Friend Function FblnEsValidoDescuento(ashrIdItemFac As Short, adecValorDscto As Decimal,
            aenuTipoDescuento As EnuTipoDescuentoDef, ByRef astrMens As String,
            ByRef aenuSevNot As EnuSeveridadNot) As Boolean
        Dim lblnEsValido = True, ldecBase = 0D, ldblTasa = 0.0
        Dim lobjItemFac As ClsItemFactura = ColItemsFactura(ashrIdItemFac.ToString)
        If aenuTipoDescuento >= EnuTipoDescuentoDef.EnuReteFuente AndAlso aenuTipoDescuento <=
                        EnuTipoDescuentoDef.EnuReteCree Then
            lblnEsValido = Not lobjItemFac.FblnRetencionAplicada(aenuTipoDescuento)
            astrMens = "Este tipo de Retención ya fue aplicada!"
            aenuSevNot = EnuSeveridadNot.EnuDatoInvalido
        End If
        If lblnEsValido Then
            Select Case aenuTipoDescuento
                Case EnuTipoDescuentoDef.EnuDsctoCapital
                    lblnEsValido = adecValorDscto <= (lobjItemFac.FdecDeudaServicioTotal -
                                    lobjItemFac.FdecDeudaIva)
                    If Not lblnEsValido Then
                        astrMens = "El Valor ingresado es mayor a la Deuda de Capital!"
                        aenuSevNot = EnuSeveridadNot.EnuDatoInvalido
                    End If
                Case EnuTipoDescuentoDef.EnuDsctoIntMora
                    lblnEsValido = adecValorDscto <= lobjItemFac.FdecDeudaIntMora
                    If Not lblnEsValido Then
                        astrMens = "El Valor ingresado es mayor a la Deuda de Intereses de Mora!"
                        aenuSevNot = EnuSeveridadNot.EnuDatoInvalido
                    End If
                Case EnuTipoDescuentoDef.EnuReteFuente, EnuTipoDescuentoDef.EnuReteIca
                    lblnEsValido = adecValorDscto <= (lobjItemFac.FdecDeudaServicioTotal -
                                    lobjItemFac.FdecDeudaIva)
                    If lblnEsValido Then
                        If lobjItemFac.FdecValorRetencion(aenuTipoDescuento, ldecBase, ldblTasa) <>
                                        adecValorDscto Then
                            astrMens = "El valor de la Retención ingresado no concuerda con los " &
                                            "Parámetros del Servicio"
                            aenuSevNot = EnuSeveridadNot.EnuAdvertencia
                        End If
                    Else
                        astrMens = "El Valor ingresado es mayor a la Deuda de Capital!"
                        aenuSevNot = EnuSeveridadNot.EnuDatoInvalido
                    End If
                Case EnuTipoDescuentoDef.EnuReteIva
                    lblnEsValido = (adecValorDscto < lobjItemFac.FdecDeudaServicioTotal)
                    If lblnEsValido Then
                        If lobjItemFac.FdecValorRetencion(aenuTipoDescuento, ldecBase, ldblTasa) <>
                                        adecValorDscto Then
                            astrMens = "El valor de la Retención ingresado no concuerda con los " &
                                            "Parámetros del Servicio"
                            aenuSevNot = EnuSeveridadNot.EnuAdvertencia
                        End If
                    Else
                        astrMens = "El Valor ingresado es mayor a la Deuda de Capital!"
                        aenuSevNot = EnuSeveridadNot.EnuDatoInvalido
                    End If
                Case EnuTipoDescuentoDef.EnuCancelaIva
                    lblnEsValido = (adecValorDscto = lobjItemFac.FdecIvaServicio +
                            lobjItemFac.FdecBaseIvaInt)
                    If Not lblnEsValido Then
                        astrMens = "Solo se puede descontar el total del IVA cuando se está " &
                                "anulando el Servicio!"
                        aenuSevNot = EnuSeveridadNot.EnuDatoInvalido
                    End If
            End Select
        End If
        Return lblnEsValido
    End Function
    ''' <summary>
    ''' Aplica un crédito originado en un recibo de caja.
    ''' </summary>
    ''' <param name="aobjItemRec">El Item del recibo de caja que tiene la informacion del crédito.</param>
    Friend Sub SApliqueCreditoRC(aobjItemRec As ClsItemRecCaja)
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
            EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
        End If
        Dim lenuTipoNov As EnuTipoNov = aobjItemRec.FenuTipoNovedad()
        Dim lblnACapital = Not (lenuTipoNov = EnuTipoNov.EnuCrDctoInt OrElse
                          lenuTipoNov = EnuTipoNov.EnuCrPagoInt)
        If lblnACapital Then
            Dim lblnEsDsctoPP = aobjItemRec.ObjIdTipoItemRecByt.ObjValorPro = EnuTipoItemRecCajaDef.EnuDsctoPP
            If lblnEsDsctoPP Then
                SApliqueDsctoPP(aobjItemRec)
            Else
                SApliqueCrRecCap(aobjItemRec)
            End If
        Else
            SApliqueCrRecIntMora(aobjItemRec)
        End If
    End Sub
    Private Sub SApliqueCrRecIntMora(aobjItemRec As ClsItemRecCaja) 'Ok
        Dim ldecVlrAplicar As Decimal, ldecVlrTotalAplicado = 0D
        Dim ldecVlrCr As Decimal = aobjItemRec.ObjValor_ItemRecDec.ObjValorPro
        Dim ldecVlrPorAplicar = ldecVlrCr
        Dim lenuTipoNov As EnuTipoNov = aobjItemRec.FenuTipoNovedad()
        Dim lobjRecCaja As ClsReciboCaja = aobjItemRec.ObjPadre
        Dim ldtmFechaCr As Date = lobjRecCaja.ObjFechaRecDtm.ObjValorPro
        Dim lstrIdCtaDb As String = aobjItemRec.ObjIdCuentaDb_ItemRecStr.ObjValorPro
        Dim ldecBase As Decimal, ldblFactor As Double
        For Each lobjItemFra As ClsItemFactura In McolItemsFactura
            ldecVlrAplicar = lobjItemFra.FdecDeudaIntMora
            If ldecVlrAplicar > 0 Then
                If ldecVlrPorAplicar < ldecVlrAplicar Then
                    ldecVlrAplicar = ldecVlrPorAplicar
                End If
                ldecBase = aobjItemRec.ObjBaseDsctoDec.ObjValorPro
                ldblFactor = aobjItemRec.ObjTasaDsctoDbl.ObjValorPro
                ldecVlrTotalAplicado += ldecVlrAplicar
                If ldecVlrAplicar > 0 Then
                    lobjItemFra.SApliqueCredito(lenuTipoNov, ldtmFechaCr, lstrIdCtaDb,
                            aobjItemRec.ObjIdItemRecCajaShr.ObjValorPro,
                            lobjRecCaja.ObjPrefijo_RecStr.ObjValorPro,
                            lobjRecCaja.ObjIdRecCajaEnt.ObjValorPro,
                            EnuTipoDocOri.EnuReciboCaja, ldecVlrAplicar,
                            ldecBase, ldblFactor)
                    ldecVlrPorAplicar -= ldecVlrAplicar
                End If
            End If
            If ldecVlrPorAplicar = 0 Then Exit For
        Next
        ObjCreditos_FactDec.ObjValorPro += ldecVlrTotalAplicado
        If ldecVlrCr <> ldecVlrTotalAplicado Then
            Throw New ErrorInesperadoPanLException("Valor aplicado difiere de valor del crédito!")
        End If
        SVerifiqueCancelacion(ldtmFechaCr)
    End Sub
    Private Sub SApliqueCrRecCap(aobjItemRec As ClsItemRecCaja)
        Dim ldecVlrAplicar As Decimal, ldecVlrTotalAplicado = 0D
        Dim ldecVlrCr As Decimal = aobjItemRec.ObjValor_ItemRecDec.ObjValorPro
        Dim ldecVlrPorAplicar = ldecVlrCr
        Dim lenuTipoNov As EnuTipoNov = aobjItemRec.FenuTipoNovedad()
        Dim lobjRecCaja As ClsReciboCaja = aobjItemRec.ObjPadre
        Dim ldtmFechaCr As Date = lobjRecCaja.ObjFechaRecDtm.ObjValorPro
        Dim lstrIdCtaDb As String = aobjItemRec.ObjIdCuentaDb_ItemRecStr.ObjValorPro
        Dim ldecBase As Decimal, ldblFactor As Double
        Dim lobjPadreItemRC As ClsReciboCaja = aobjItemRec.ObjPadre
        Dim lstrServicio As String = lobjPadreItemRC.ObjServicios_RecStr.ObjValorPro
        Dim lshrIdItemFac As Short = aobjItemRec.ObjIdItemFac_ItemRecShr.ObjValorPro
        Dim lobjItemFra As ClsItemFactura = ColItemsFactura(lshrIdItemFac.ToString)
        ldecVlrAplicar = lobjItemFra.FdecDeudaServicioTotal
        If ldecVlrAplicar > 0 Then
            If ldecVlrAplicar > ldecVlrPorAplicar Then
                ldecVlrAplicar = ldecVlrPorAplicar
            End If
            ldecVlrTotalAplicado += ldecVlrAplicar
            ldecBase = aobjItemRec.ObjBaseDsctoDec.ObjValorPro
            ldblFactor = aobjItemRec.ObjTasaDsctoDbl.ObjValorPro
            If ldecVlrAplicar > 0 Then
                lobjItemFra.SApliqueCredito(lenuTipoNov, ldtmFechaCr, lstrIdCtaDb,
                        aobjItemRec.ObjIdItemRecCajaShr.ObjValorPro, lobjRecCaja.ObjPrefijo_RecStr.ObjValorPro,
                        lobjRecCaja.ObjIdRecCajaEnt.ObjValorPro,
                        EnuTipoDocOri.EnuReciboCaja, ldecVlrAplicar,
                        ldecBase, ldblFactor)
            End If
        End If
        ObjCreditos_FactDec.ObjValorPro += ldecVlrTotalAplicado
        If ldecVlrCr <> ldecVlrTotalAplicado Then
            Throw New ErrorInesperadoPanLException("Valor aplicado difiere de valor del crédito!")
        End If
        SVerifiqueCancelacion(ldtmFechaCr)
    End Sub
    Private Sub SApliqueDsctoPP(aobjItemRec As ClsItemRecCaja)
        Dim lenuTipoNov As EnuTipoNov = aobjItemRec.FenuTipoNovedad()
        Dim lobjRecCaja As ClsReciboCaja = aobjItemRec.ObjPadre
        Dim ldtmFechaCr As Date = lobjRecCaja.ObjFechaRecDtm.ObjValorPro
        Dim lstrIdCtaDb As String = aobjItemRec.ObjIdCuentaDb_ItemRecStr.ObjValorPro
        Dim ldecVlrAplicar As Decimal, ldecVlrTotalAplicado = 0D
        Dim ldecVlrCr As Decimal = aobjItemRec.ObjValor_ItemRecDec.ObjValorPro
        Dim ldecValorPorAplicar = ldecVlrCr, ldecBase As Decimal, ldblFactor As Double
        For Each lobjItemFra As ClsItemFactura In McolItemsFactura
            If lobjItemFra.FdecDeudaServicioTotal > 0 AndAlso
                    lobjItemFra.ObjServicio.BlnEsCuotaAdministracion AndAlso
                    Not lobjItemFra.ObjServicio.ObjEsAjusteBln.ObjValorPro Then
                If lobjItemFra.FdecDeudaServicioTotal >= ldecValorPorAplicar Then
                    ldecVlrAplicar = ldecValorPorAplicar
                Else
                    ldecVlrAplicar = lobjItemFra.FdecDeudaServicioTotal
                End If
                If ldecVlrAplicar > 0 Then
                    ldecBase = aobjItemRec.ObjBaseDsctoDec.ObjValorPro
                    ldblFactor = aobjItemRec.ObjTasaDsctoDbl.ObjValorPro
                    lobjItemFra.SApliqueCredito(lenuTipoNov, ldtmFechaCr, lstrIdCtaDb,
                            aobjItemRec.ObjIdItemRecCajaShr.ObjValorPro,
                            lobjRecCaja.ObjPrefijo_RecStr.ObjValorPro,
                            lobjRecCaja.ObjIdRecCajaEnt.ObjValorPro,
                            EnuTipoDocOri.EnuReciboCaja, ldecVlrAplicar,
                            ldecBase, ldblFactor)
                    ldecValorPorAplicar -= ldecVlrAplicar
                End If
                ldecVlrTotalAplicado += ldecVlrAplicar
                If ldecVlrTotalAplicado = ldecVlrCr Then Exit For
            End If
        Next
        ObjCreditos_FactDec.ObjValorPro += ldecVlrTotalAplicado
        If ldecVlrCr <> ldecVlrTotalAplicado Then
            Throw New ErrorInesperadoPanLException("Valor aplicado difiere de valor del crédito!")
        End If
        SVerifiqueCancelacion(ldtmFechaCr)
    End Sub
    Friend Sub SApliqueCreditoRC(aobjItemRec As ClsItemRecCaja, astrIdPredio As String)
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
            SModifique()
        End If
        Dim lenuTipoNov As EnuTipoNov = aobjItemRec.FenuTipoNovedad()
        McolItemsFactura = ColItemsFactura
        Dim lblnACapital = Not (lenuTipoNov = EnuTipoNov.EnuCrDctoInt OrElse
                          lenuTipoNov = EnuTipoNov.EnuCrPagoInt)
        If lblnACapital Then
            SApliqueCrRecCap(aobjItemRec, astrIdPredio)
        Else
            SApliqueCrRecIntMora(aobjItemRec, astrIdPredio)
        End If
    End Sub
    Private Sub SApliqueCrRecIntMora(aobjItemRec As ClsItemRecCaja, astrIdPredio As String) 'Ok
        Dim ldecVlrAplicar As Decimal, ldecVlrTotalAplicado = 0D
        Dim ldecVlrCr As Decimal = aobjItemRec.ObjValor_ItemRecDec.ObjValorPro
        Dim ldecVlrPorAplicar = ldecVlrCr
        Dim lenuTipoNov As EnuTipoNov = aobjItemRec.FenuTipoNovedad()
        Dim lobjRecCaja As ClsReciboCaja = aobjItemRec.ObjPadre
        Dim ldtmFechaCr As Date = lobjRecCaja.ObjFechaRecDtm.ObjValorPro
        Dim lstrIdCtaDb As String = aobjItemRec.ObjIdCuentaDb_ItemRecStr.ObjValorPro
        Dim ldecBase As Decimal, ldblFactor As Double
        For Each lobjItemFra As ClsItemFactura In McolItemsFactura
            If lobjItemFra.ObjIdPredio_ItemFactStr.ObjValorPro = astrIdPredio Then
                ldecVlrAplicar = lobjItemFra.FdecDeudaIntMora
                If ldecVlrAplicar > ldecVlrPorAplicar Then
                    ldecVlrAplicar = ldecVlrPorAplicar
                End If
                If ldecVlrAplicar > 0 Then
                    ldecVlrTotalAplicado += ldecVlrAplicar
                    ldecBase = aobjItemRec.ObjBaseDsctoDec.ObjValorPro
                    ldblFactor = aobjItemRec.ObjTasaDsctoDbl.ObjValorPro
                    lobjItemFra.SApliqueCredito(lenuTipoNov, ldtmFechaCr, lstrIdCtaDb,
                            aobjItemRec.ObjIdItemRecCajaShr.ObjValorPro, lobjRecCaja.ObjPrefijo_RecStr.ObjValorPro,
                            lobjRecCaja.ObjIdRecCajaEnt.ObjValorPro,
                            EnuTipoDocOri.EnuReciboCaja, ldecVlrAplicar,
                            ldecBase, ldblFactor)
                    ldecVlrPorAplicar -= ldecVlrAplicar
                End If
            End If
        Next
        ObjCreditos_FactDec.ObjValorPro += ldecVlrTotalAplicado
        If ldecVlrCr <> ldecVlrTotalAplicado Then
            Throw New ErrorInesperadoPanLException("Valor aplicado difiere de valor del crédito!")
        End If
        SVerifiqueCancelacion(ldtmFechaCr)
    End Sub
    Private Sub SApliqueCrRecCap(aobjItemRec As ClsItemRecCaja, astrIdPredio As String)
        Dim ldecVlrAplicar As Decimal, ldecVlrTotalAplicado = 0D
        Dim ldecVlrCr As Decimal = aobjItemRec.ObjValor_ItemRecDec.ObjValorPro
        Dim ldecVlrPorAplicar = ldecVlrCr
        Dim lenuTipoNov As EnuTipoNov = aobjItemRec.FenuTipoNovedad()
        Dim lobjRecCaja As ClsReciboCaja = aobjItemRec.ObjPadre
        Dim ldtmFechaCr As Date = lobjRecCaja.ObjFechaRecDtm.ObjValorPro
        Dim lstrIdCtaDb As String = aobjItemRec.ObjIdCuentaDb_ItemRecStr.ObjValorPro
        Dim ldecBase As Decimal, ldblFactor As Double
        For Each lobjItemFra As ClsItemFactura In McolItemsFactura
            If lobjItemFra.ObjIdPredio_ItemFactStr.ObjValorPro = astrIdPredio Then
                ldecVlrAplicar = lobjItemFra.FdecDeudaServicioTotal
                If ldecVlrAplicar > ldecVlrPorAplicar Then
                    ldecVlrAplicar = ldecVlrPorAplicar
                End If
                If ldecVlrAplicar > 0 Then
                    ldecBase = aobjItemRec.ObjBaseDsctoDec.ObjValorPro
                    ldblFactor = aobjItemRec.ObjTasaDsctoDbl.ObjValorPro
                    lobjItemFra.SApliqueCredito(lenuTipoNov, ldtmFechaCr, lstrIdCtaDb,
                            aobjItemRec.ObjIdItemRecCajaShr.ObjValorPro, lobjRecCaja.ObjPrefijo_RecStr.ObjValorPro,
                            lobjRecCaja.ObjIdRecCajaEnt.ObjValorPro,
                            EnuTipoDocOri.EnuReciboCaja, ldecVlrAplicar,
                            ldecBase, ldblFactor)
                    ldecVlrPorAplicar -= ldecVlrAplicar
                    ldecVlrTotalAplicado += ldecVlrAplicar
                End If
            End If
        Next
        ObjCreditos_FactDec.ObjValorPro += ldecVlrTotalAplicado
        If ldecVlrCr <> ldecVlrTotalAplicado Then
            Throw New ErrorInesperadoPanLException("Valor aplicado difiere de valor del crédito!")
        End If
        SVerifiqueCancelacion(ldtmFechaCr)
    End Sub
    ''' <summary>
    ''' Aplica un crédito originado en una nota de aplicación de anticipo cuando hay dscto por pronto pago
    ''' </summary>
    ''' <param name="aobjItemNotaCon">El Item de la nota contable que tiene la informacion del crédito.</param>
    Friend Sub SApliqueCreditoNotaCon(aobjItemNotaCon As ClsItemNotaCon)
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
            EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
        End If
        Dim lenuTipoNov As EnuTipoNov = EnuTipoNov.EnuCrDctoCap
        Dim ldecVlrAplicar As Decimal, ldecVlrTotalAplicado = 0D
        Dim ldecVlrCr As Decimal = aobjItemNotaCon.ObjValor_ItemNotaConDec.ObjValorPro
        Dim ldecVlrPorAplicar = ldecVlrCr
        Dim lobjNotaCon As ClsNotaCon = aobjItemNotaCon.ObjPadre
        Dim ldtmFechaCr As Date = lobjNotaCon.ObjFecha_NotaConDtm.ObjValorPro
        Dim lstrIdCtaDb As String = GobjParametros.ObjIdCtaDescuentosPPStr.ObjValorPro
        Dim ldecBase = 0D, ldblFactor = 0.0
        For Each lobjItemFra As ClsItemFactura In ColItemsFactura
            ldecVlrAplicar = lobjItemFra.FdecDeudaServicioTotal
            If ldecVlrAplicar > ldecVlrPorAplicar Then
                ldecVlrAplicar = ldecVlrPorAplicar
            End If
            If ldecVlrAplicar > 0 Then
                ldecVlrTotalAplicado += ldecVlrAplicar
                lobjItemFra.SApliqueCredito(lenuTipoNov, ldtmFechaCr, lstrIdCtaDb,
                        aobjItemNotaCon.ObjIdItemNotaConShr.ObjValorPro,
                        lobjNotaCon.ObjPrefijo_NotaConStr.ObjValorPro,
                        lobjNotaCon.ObjIdNotaConEnt.ObjValorPro,
                        EnuTipoDocOri.EnuNotaCon, ldecVlrAplicar,
                        ldecBase, ldblFactor)
                ldecVlrPorAplicar -= ldecVlrAplicar
            End If
        Next
        ObjCreditos_FactDec.ObjValorPro += ldecVlrTotalAplicado
        If ldecVlrCr <> ldecVlrTotalAplicado Then
            Throw New ErrorInesperadoPanLException("Valor aplicado difiere de valor del crédito!")
        End If
        SVerifiqueCancelacion(ldtmFechaCr)
    End Sub
    Friend Sub SApliqueDscto(aobjItemRec As ClsItemRecCaja)
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
            EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
        End If
        With aobjItemRec
            Dim lobjRec As ClsReciboCaja = aobjItemRec.ObjPadre
            Dim lobjItemFact As ClsItemFactura = ColItemsFactura(.ObjIdItemFac_ItemRecShr.ToString)
            Dim lenuTipoNov As EnuTipoNov = aobjItemRec.FenuTipoNovedad()
            Dim ldecBase As Decimal = aobjItemRec.ObjBaseDsctoDec.ObjValorPro
            Dim ldblFactor As Double = aobjItemRec.ObjTasaDsctoDbl.ObjValorPro
            lobjItemFact.SApliqueCredito(lenuTipoNov, lobjRec.ObjFechaRecDtm.ObjValorPro,
                               aobjItemRec.ObjIdCuentaDb_ItemRecStr.ObjValorPro, aobjItemRec.ObjIdItemRecCajaShr.ObjValorPro,
                               lobjRec.ObjPrefijo_RecStr.ObjValorPro, lobjRec.ObjIdRecCajaEnt.ObjValorPro,
                               EnuTipoDocOri.EnuReciboCaja, aobjItemRec.ObjValor_ItemRecDec.ObjValorPro,
                               ldecBase, ldblFactor)
            ObjCreditos_FactDec.ObjValorPro += .ObjValor_ItemRecDec.ObjValorPro
            SVerifiqueCancelacion(lobjRec.ObjFechaRecDtm.ObjValorPro)
        End With
    End Sub
    Private Sub SVerifiqueCancelacion(adtmFechaCredito As Date)
        If ObjCreditos_FactDec.ObjValorPro > ObjDebitos_FactDec.ObjValorPro Then
            Throw New ErrorInesperadoPanLException("Creditos mayores a los debitos en la factura!")
        End If
        If ObjCreditos_FactDec.ObjValorPro = ObjDebitos_FactDec.ObjValorPro Then
            ObjFechaCancelacion_FactDtm.ObjValorPro = adtmFechaCredito
        End If
    End Sub
    ''' <summary>
    ''' Aplica el anticipo generado en un Recibo de Caja
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub SApliqueAnticipo(aobjItemNotaCon As ClsItemNotaCon)
        Dim lobjNotaCon As ClsNotaCon = aobjItemNotaCon.ObjPadre
        Dim ldtmFechaTransaccion As Date = lobjNotaCon.ObjFecha_NotaConDtm.ObjValorPro
        Dim lstrPrefNotaCon As String = aobjItemNotaCon.ObjPrefijoNotaCon_ItemNotaConStr.ObjValorPro
        Dim lentIdNotaCon As Integer = aobjItemNotaCon.ObjIdNotaCon_ItemNotaConEnt.ObjValorPro
        Dim lshrIdItemFac As Short = aobjItemNotaCon.ObjIdItemFac_ItemNotaConShr.ObjValorPro
        Dim ldecValor As Decimal = aobjItemNotaCon.ObjValor_ItemNotaConDec.ObjValorPro
        Dim lobjItemFac As ClsItemFactura = ColItemsFactura(lshrIdItemFac.ToString)
        Dim lenuTipoItemNotaCon = EnuTipoItemNotaConDef.None, lstrIdCtaDb = String.Empty
        Dim ldecVlrBase = 0D, ldblTasa As Double
        Dim lshrIdItemNotCon As Short = aobjItemNotaCon.ObjIdItemNotaConShr.ObjValorPro
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
            EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
        End If
        With lobjItemFac
            If .EnuEstadoActualizacion <> EnuEstadoObjetoDef.EnuModificando Then
                .EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
            End If
            lenuTipoItemNotaCon = aobjItemNotaCon.ObjIdTipoItemNotaConByt.ObjValorPro
            Select Case lenuTipoItemNotaCon
                Case EnuTipoItemNotaConDef.EnuAplicaAntCap
                    lstrIdCtaDb = GobjParametros.ObjIdCtaAnticiposRecibidosStr.ObjValorPro
                    .SApliqueCredito(EnuTipoNov.EnuCrAnApCap, ldtmFechaTransaccion, lstrIdCtaDb,
                            lshrIdItemNotCon, lstrPrefNotaCon, lentIdNotaCon,
                            EnuTipoDocOri.EnuNotaCon, ldecValor, ldecVlrBase, ldblTasa)
                Case EnuTipoItemNotaConDef.EnuAplicaAntInt
                    lstrIdCtaDb = GobjParametros.ObjIdCtaAnticiposRecibidosStr.ObjValorPro
                    .SApliqueCredito(EnuTipoNov.EnuCrAnApInt, ldtmFechaTransaccion, lstrIdCtaDb,
                            lshrIdItemNotCon, lstrPrefNotaCon, lentIdNotaCon,
                            EnuTipoDocOri.EnuNotaCon, ldecValor, ldecVlrBase, ldblTasa)
                Case EnuTipoItemNotaConDef.EnuDsctoPP
                    lstrIdCtaDb = GobjParametros.ObjIdCtaDescuentosPPStr.ObjValorPro
                    ldecVlrBase = .FdecValorServicio
                    ldblTasa = aobjItemNotaCon.ObjValor_ItemNotaConDec.ObjValorPro / ldecVlrBase
                    .SApliqueCredito(EnuTipoNov.EnuCrDctoCap, ldtmFechaTransaccion, lstrIdCtaDb,
                            lshrIdItemNotCon, lstrPrefNotaCon, lentIdNotaCon,
                            EnuTipoDocOri.EnuNotaCon, ldecValor, ldecVlrBase, ldblTasa)
                Case EnuTipoItemNotaConDef.EnuReteFuente
                    lstrIdCtaDb = GobjParametros.ObjIdCtaReteFuenteStr.ObjValorPro
                    ldecVlrBase = .FdecValorServicio
                    ldblTasa = ldecValor / ldecVlrBase
                    .SApliqueCredito(EnuTipoNov.EnuCrRetFte, ldtmFechaTransaccion, lstrIdCtaDb,
                            lshrIdItemNotCon, lstrPrefNotaCon, lentIdNotaCon,
                            EnuTipoDocOri.EnuNotaCon, ldecValor, ldecVlrBase, ldblTasa)
                Case EnuTipoItemNotaConDef.EnuReteIca
                    lstrIdCtaDb = GobjParametros.ObjIdCtaReteIcaStr.ObjValorPro
                    ldecVlrBase = .FdecValorServicio
                    ldblTasa = ldecValor / ldecVlrBase
                    .SApliqueCredito(EnuTipoNov.EnuCrRetIca, ldtmFechaTransaccion, lstrIdCtaDb,
                            lshrIdItemNotCon, lstrPrefNotaCon, lentIdNotaCon,
                            EnuTipoDocOri.EnuNotaCon, ldecValor, ldecVlrBase, ldblTasa)
                Case EnuTipoItemNotaConDef.EnuReteIva
                    lstrIdCtaDb = GobjParametros.ObjIdCtaReteIvaStr.ObjValorPro
                    ldecVlrBase = .FdecValorServicio
                    ldblTasa = ldecValor / ldecVlrBase
                    .SApliqueCredito(EnuTipoNov.EnuCrRetIva, ldtmFechaTransaccion, lstrIdCtaDb,
                            lshrIdItemNotCon, lstrPrefNotaCon, lentIdNotaCon,
                            EnuTipoDocOri.EnuNotaCon, ldecValor, ldecVlrBase, ldblTasa)
            End Select
            ObjCreditos_FactDec.ObjValorPro += ldecValor
            Dim ldecDeuAdmi As Decimal
            For Each lobjItFac As ClsItemFactura In ColItemsFactura
                If lobjItemFac.ObjServicio.ObjIdTipoServicioByt.ObjValorPro =
                        EnuTipoServicio.EnuAnual Then
                    ldecDeuAdmi += lobjItemFac.DecDeuda
                End If
            Next
            If ldecDeuAdmi = 0 Then
                ObjDctoProntoPago_FacDec.ObjValorPro = 0
            End If
            SVerifiqueCancelacion(ldtmFechaTransaccion)
        End With
    End Sub
    Friend Sub SApliqueNotaCr(aobjItemNotaCr As ClsItemNotaCr)
        Dim ldecVlrCr As Decimal = aobjItemNotaCr.ObjValor_ItemNotaCrDec.ObjValorPro
        Dim lobjItemFac As ClsItemFactura = ColItemsFactura(
                aobjItemNotaCr.ObjIdItemFac_ItemNotaCrShr.ToString())
        Dim lenuTipoDscto As EnuTipoDescuentoDef =
                aobjItemNotaCr.ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro
        If lenuTipoDscto = EnuTipoDescuentoDef.EnuDsctoIntMora Then
            ' AVV Revisar
            'If ldecVlrCr > lobjItemFac.FdecDeudaIntMora - lobjItemFac.FdecDeudaIvaInt Then
            '    Throw New ErrorInesperadoPanLException("Valor Item NotaCr mayor a la deuda por Mora")
            'End If
        ElseIf lenuTipoDscto = EnuTipoDescuentoDef.EnuDsctoCapital Then
            If ldecVlrCr > (lobjItemFac.FdecDeudaServicioTotal) Then
                Throw New ErrorInesperadoPanLException("Valor Item NotaCr mayor a la deuda de Capital")
            End If
        ElseIf lenuTipoDscto = EnuTipoDescuentoDef.EnuCancelaIva Then
            If ldecVlrCr <> lobjItemFac.FdecDeudaIva Then
                Throw New ErrorInesperadoPanLException("Valor Item NotaCr difiere de la deuda del IVA")
            End If
        End If
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
            SModifique()
        End If
        SApliqueDscto(aobjItemNotaCr)
    End Sub
    Private Sub SApliqueDscto(aobjItemNotaCr As ClsItemNotaCr)
        Dim lshrIdItemFac As Short = aobjItemNotaCr.ObjIdItemFac_ItemNotaCrShr.ObjValorPro
        Dim ldecVlrCr As Decimal = aobjItemNotaCr.ObjValor_ItemNotaCrDec.ObjValorPro
        Dim lenuTipoNov As EnuTipoNov = aobjItemNotaCr.FenuTipoNovedad()
        Dim lobjNotaCr As ClsNotaCr = aobjItemNotaCr.ObjPadre
        Dim lobjItemFac As ClsItemFactura = ColItemsFactura(lshrIdItemFac.ToString)
        Dim ldecBase = aobjItemNotaCr.ObjBaseDscto_NotaCrDec.ObjValorPro
        Dim ldblFactor = aobjItemNotaCr.ObjTasaDscto_ItemNotaCrDbl.ObjValorPro
        Dim lstrIdCtaDb As String = ClsOrionCop.FstrIdCtaDbDscto(
                aobjItemNotaCr.ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro)
        SApliqueCr(lobjItemFac, aobjItemNotaCr, ldecVlrCr, lenuTipoNov, lstrIdCtaDb)
        ObjCreditos_FactDec.ObjValorPro += ldecVlrCr
        Dim ldtmFechaCr As Date = lobjNotaCr.ObjFecha_NotaCrDtm.ObjValorPro
        SVerifiqueCancelacion(ldtmFechaCr)
    End Sub
    Private Sub SApliqueCr(aobjItemFac As ClsItemFactura, aobjItemNotaCr As ClsItemNotaCr,
            adecValorCr As Decimal, aenuTipoNov As EnuTipoNov, astrIdCtaDb As String)
        Dim lobjNotaCr As ClsNotaCr = aobjItemNotaCr.ObjPadre
        Dim ldtmFechaCr As Date = lobjNotaCr.ObjFecha_NotaCrDtm.ObjValorPro
        Dim ldecBase = aobjItemNotaCr.ObjBaseDscto_NotaCrDec.ObjValorPro
        Dim ldblFactor = aobjItemNotaCr.ObjTasaDscto_ItemNotaCrDbl.ObjValorPro
        aobjItemFac.SApliqueCredito(aenuTipoNov, ldtmFechaCr, astrIdCtaDb,
                        aobjItemNotaCr.ObjIdItemNotaCrShr.ObjValorPro,
                        lobjNotaCr.ObjPrefijo_NotaCrStr.ObjValorPro,
                        lobjNotaCr.ObjIdNotaCrEnt.ObjValorPro,
                        EnuTipoDocOri.EnuNotaCr, adecValorCr,
                        ldecBase, ldblFactor)
    End Sub
    Friend Sub SAnuleFactura(aobjNotaCr As ClsNotaCr)
        For Each lobjItemNCr As ClsItemNotaCr In aobjNotaCr.ColItemsNotaCr
            Dim lobjItemFac As ClsItemFactura = ColItemsFactura(
                    lobjItemNCr.ObjIdItemFac_ItemNotaCrShr.ToString)
            lobjItemFac.SAnuleNovedadesItem(lobjItemNCr)
        Next
    End Sub
    Friend Sub SAnuleItemNotaCr(aobjItemNotaCr As ClsItemNotaCr)
        Dim lstrPrefDocOrigen As String = aobjItemNotaCr.ObjPrefijo_ItemNotaCrStr.ObjValorPro
        Dim lentIdDocOrigen As Integer = aobjItemNotaCr.ObjIdNotaCr_ItemNotaCrEnt.ObjValorPro
        Dim lshrIdItemDocOrigen As Short = aobjItemNotaCr.ObjIdItemNotaCrShr.ObjValorPro
        Dim lshrIdItemFac As Short
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
            EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
        End If
        Dim lcolNovedadesItemNotaCr As Collection = aobjItemNotaCr.ColNovedades
        Dim lobjItemFac As ClsItemFactura
        Dim lobjNotaCr As ClsNotaCr = aobjItemNotaCr.ObjPadre
        Dim ldtmFchaAnu As Date = lobjNotaCr.ObjFechaAnulacion_NotaCrDtm.ObjValorPro
        ldtmFchaAnu = DateSerial(ldtmFchaAnu.Year, ldtmFchaAnu.Month, ldtmFchaAnu.Day)
        McolItemsFactura = ColItemsFactura
        For Each lobjNov As ClsNovedad In lcolNovedadesItemNotaCr
            lshrIdItemFac = lobjNov.ObjIdItemFact_NovShr.ObjValorPro
            lobjItemFac = McolItemsFactura(lshrIdItemFac.ToString)
            If lobjItemFac.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                lobjItemFac.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
            End If
            lobjItemFac.SReverseNovedad(lobjNov, ldtmFchaAnu, EnuTipoDocOri.EnuNotaCr,
                    lstrPrefDocOrigen, lentIdDocOrigen, lshrIdItemDocOrigen)
        Next
        ObjDebitos_FactDec.ObjValorPro += aobjItemNotaCr.ObjValor_ItemNotaCrDec.ObjValorPro
        ObjFechaCancelacion_FactDtm.ObjValorPro = GCDTMFECHANULA
    End Sub
#End Region
#Region "Retenciones y descuento por pronto pago"
    ''' <summary>
    ''' Devuelve el valor del descuento por pronto pago a que se tiene derecho si se paga antes o 
    ''' en la fecha indicada
    ''' </summary>
    ''' <remarks>Se calcula al momento de generar la factura y se guarda como una 
    ''' propiedad de la factura</remarks>
    Private Function FdecDsctoPPPosible() As Decimal
        Dim ldecDsctoPP = 0D
        For Each lobjItemFact As ClsItemFactura In ColItemsFactura
            If lobjItemFact.ObjServicio.BlnEsCuotaAdministracion Then
                ldecDsctoPP += lobjItemFact.FdecDsctoPPPosible()
            End If
        Next
        Return ldecDsctoPP
    End Function
    Friend Function FdecDsctoPPAAplicar(adtmFechaAplicacion As Date)
        Dim ldecDsctoPP = 0D
        If adtmFechaAplicacion <= ObjFechaDctoProntoPagoDtm.ObjValorPro Then
            ldecDsctoPP = ObjDctoProntoPago_FacDec.ObjValorPro
        End If
        Return ldecDsctoPP
    End Function
    Private Function FdtmFechaDsctoPP() As Date
        Dim ldtmFechaDsctoPP = GCDTMFECHANULA
        Dim ldtmFechaFac As Date = ObjFechaFacturaDtm.ObjValorPro
        If GobjParametros.ObjAnoActual.ObjAplicaDsctoPPBln.ObjValorPro Then
            ldtmFechaDsctoPP = ldtmFechaFac.AddDays(
                    GobjParametros.ObjAnoActual.ObjDiasParaDsctoPPShr.ObjValorPro - 1)
        End If
        Return ldtmFechaDsctoPP
    End Function
#End Region
#Region "Anticipos Aplicados"
    Friend Function FdecAnticipoAplicado() As Decimal
        Dim ldecAntApl = 0D, ldecAnt As Decimal
        Dim ldtbNov = DtbNovedadesFact
        Dim lstrFiltro = ClsIdTipoDocOrigen_AntByt.SstrNombreCampoBd & " = " &
                EnuTipoDocOri.EnuNotaCon
        Dim ldrwNovs As DataRow() = ldtbNov.Select(lstrFiltro)
        For Each ldrwNov As DataRow In ldrwNovs
            ldecAnt = ClsPanorama.FobjValorCampo(ldrwNov(
                    ClsValor_NovDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
            ldecAntApl += ldecAnt
        Next
        Return ldecAntApl
    End Function
#End Region
#Region "Estado Cuenta"
    Friend ReadOnly Property ObjEstadoCuenta As ClsEstadoCuenta
        Get
            If IsNothing(MobjEstadoCuenta) AndAlso BlnExiste Then
                SCargueDtbEstadoCuenta()
                Dim ldrwEstadosCuentas As DataRow() = MdtbEstadoCuenta.Select
                If ldrwEstadosCuentas.Length > 0 Then
                    Dim ldrwEstadoCuenta As DataRow = ldrwEstadosCuentas(0)
                    MobjEstadoCuenta = New ClsEstadoCuenta(ldrwEstadoCuenta)
                    MobjEstadoCuenta.SLeaValores(True)
                End If
            End If
            Return MobjEstadoCuenta
        End Get
    End Property
    Private Sub SCargueDtbEstadoCuenta()
        If IsNothing(MdtbEstadoCuenta) Then
            Dim lstrIdFactura = ObjIdFacturaEnt.ToString
            If String.IsNullOrEmpty(lstrIdFactura) Then lstrIdFactura = "0"
            Dim lstrIndice = {{StrCampoCarpeta, "ASC"},
                              {StrCampoCentroUtil, "ASC"},
                              {ClsPrefijoFac_EstadoStr.SstrNombreCampoBd, "ASC"},
                              {ClsIdFactura_EstadoEnt.SstrNombreCampoBd, "ASC"}}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsPrefijoFac_EstadoStr.SstrNombreCampoBd & " = '" & ObjPrefijo_FactStr.ToString &
                    "' AND " & ClsIdFactura_EstadoEnt.SstrNombreCampoBd & " = " & lstrIdFactura
            Dim lstrCamposSelect() = {"*"}
                MdtbEstadoCuenta = ClsPanorama.FdtbDataTable(ClsEstadoCuenta.SstrNombreTabla, lstrCamposSelect,
                    lstrIndice, lstrFiltro)
            End If
    End Sub
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsCreditos_FactDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Creditos"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "CreditosFactura"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, Decimal.MaxValue,
                    BlnEsRequerido, EnuTipoValor.enuDecimal)
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
Friend Class ClsCUFEStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "CUFE"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Código Unico Factura Electrónica"
        HenuTipoValor = EnuTipoValor.enuString
        HshrLongitud = 300
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
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
        If Not IsNothing(ObjValorPro) Then
            Return HobjValorPro.ToString
        Else
            Return ""
        End If
    End Function
End Class
Friend Class ClsDctoProntoPago_FacDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "DctoProntoPago"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "DctoProntoPago"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = False
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        Dim lblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1,
                Decimal.MaxValue, BlnEsRequerido, HenuTipoValor)
        HblnEsValido = lblnEsValido
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsDebitos_FactDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Debitos"
    Private ReadOnly MobjPadre As ClsFactura = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Debitos"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Decimal.MaxValue,
                BlnEsRequerido, HenuTipoValor)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            End If
        Else
            If HobjValorNew = 0 Then
                HblnEsValido = MobjPadre.ObjAnuladoBln.ObjValorPro
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
Friend Class ClsEnviadaMailBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "EnviadaMail"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Enviada por Email"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
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
        Return ClsPanorama.FstrBuleanoToString(HobjValorPro)
    End Function
End Class
Friend Class ClsEsPreFacturaBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "EsPrefactura"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "EsPrefactura"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
    End Sub
    Private Sub ClsEsPreFacturaBln_evnPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        Dim lobjPadre As ClsCBObjetoPan = ObjPadre
        If TypeOf lobjPadre Is ClsFactura Then
            Dim lobjFactPadre As ClsFactura = ObjPadre
            lobjFactPadre.ObjPrefijo_FactStr.SValide()
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        Return ClsPanorama.FstrBuleanoToString(HobjValorPro)
    End Function
End Class
Friend Class ClsFechaAnulacion_FactDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaAnulacion"
    Private ReadOnly MobjPadre As ClsFactura = Nothing
    Public Sub New(aobjPadre As ClsFactura)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "FechaAnulacion_Fac"
        HenuTipoValor = EnuTipoValor.enuDateTime
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = False
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MyBase.SVaciePropiedad()
        HobjValorNew = GCDTMFECHANULA
        HobjValorPro = HobjValorNew
    End Sub
    Public Overrides Sub SValide()
        Dim ldtmFechaMin = GCDTMFECHANULA
        Dim ldtmFechaMax = Date.Today
        Select Case MobjPadre.EnuEstadoActualizacion
            Case EnuEstadoObjetoDef.enuCreando
                '
            Case EnuEstadoObjetoDef.enuModificando
                If Not ClsOrionCop.BlnFacturando Then
                    If GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro Then
                        ldtmFechaMin = Now.AddHours(-Now.Hour)
                    Else
                        If GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo < Date.Today Then
                            ldtmFechaMin = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
                        Else
                            ldtmFechaMin = Now
                        End If
                    End If
                    ldtmFechaMax = Now
                End If
            Case EnuEstadoObjetoDef.enuConsultando
                ldtmFechaMin = GCDTMFECHANULA
                ldtmFechaMax = Now
        End Select
        HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax,
                BlnEsRequerido)
    End Sub
    Private Sub EPosSetValor() Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando AndAlso BlnEsValido Then
            If HobjValorPro <> GCDTMFECHANULA Then
                MobjPadre.ObjValor_FactDec.SValide()
                MobjPadre.ObjCreditos_FactDec.SValide()
                MobjPadre.ObjDebitos_FactDec.SValide()
            End If
        End If
    End Sub
    Public Overrides Function ToString() As String
        If Not IsNothing(HobjValorPro) Then
            Return Format(HobjValorPro, GCSTRFMTFECHASIMPLE)
        Else
            Return Format(GCDTMFECHANULA, GCSTRFMTFECHASIMPLE)
        End If
    End Function
End Class
Friend Class ClsFechaCancelacion_FactDtm
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsFactura = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaCancelacion"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "FechaCancelacionFactura"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = False
    End Sub
    Public Overrides Sub SValide()
        Dim ldtmFechaMin As Date = GCDTMFECHANULA
        Dim ldtmFechaMax As Date = Date.Today
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
            If MobjPadre.ObjDebitos_FactDec.ObjValorPro = MobjPadre.ObjCreditos_FactDec.ObjValorPro Then
                Dim ldtmFechaIniPeriodoActual = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
                If ldtmFechaIniPeriodoActual >= MobjPadre.ObjFechaFacturaDtm.ObjValorPro Then
                    ldtmFechaMin = ldtmFechaIniPeriodoActual
                Else
                    ldtmFechaMin = MobjPadre.ObjFechaFacturaDtm.ObjValorPro
                End If
            End If
            If Date.Today > GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo Then
                ldtmFechaMax = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
            Else
                ldtmFechaMax = Now
            End If
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax, BlnEsRequerido)
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        Dim lobjpadre As ClsFactura = ObjPadre
        If lobjpadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
            If HobjValorPro <> GCDTMFECHANULA Then
                HblnEsValido = (lobjpadre.DecDeuda = 0)
            End If
        End If
    End Sub
    Public Overrides Function ToString() As String
        If Not IsNothing(HobjValorPro) Then
            Return Format(HobjValorPro, GCSTRFMTFECHASIMPLE)
        Else
            Return Format(GCDTMFECHANULA, GCSTRFMTFECHASIMPLE)
        End If
    End Function
End Class
Friend Class ClsFechaDctoProntoPagoDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaDctoProntoPago"
    ReadOnly MobjPadre As ClsFactura = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "FechaDctoProntoPago"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = False
    End Sub
    Public Overrides Sub SValide()
        Dim ldtmFechaMin As Date = GCDTMFECHANULA
        Dim ldtmFechaMax As Date = Date.Today
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            ldtmFechaMax = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
        ElseIf MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
            ldtmFechaMax = HobjValorOriginal
            ldtmFechaMin = HobjValorOriginal
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax, BlnEsRequerido)
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
Friend Class ClsFechaEmisionEFacStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaEmisionEFac"
    ReadOnly MobjPadre As ClsFactura = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Fecha Emision EFac"
        HenuTipoValor = EnuTipoValor.enuString
        HshrLongitud = 20
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = False
    End Sub
    Public Overrides Sub SValide()
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando AndAlso
                GobjParametros.BlnEFacAutorizado Then
            HblnEsRequerido = MobjPadre.ObjIdEstadoEDocEnt.ObjValorPro > EnuEstadoEDoc.EnuEnProceso
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 18, 20, BlnEsRequerido)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro
        End If
    End Function
End Class
Friend Class ClsFechaFacturaDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaFactura"
    Private ReadOnly MobjPadre As ClsFactura = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "FechaFactura"
        HenuTipoValor = EnuTipoValor.enuDateTime
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MyBase.SVaciePropiedad()
        HobjValorNew = GCDTMFECHANULA
        HobjValorPro = GCDTMFECHANULA
    End Sub
    Public Overrides Sub SValide()
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando OrElse
                ClsOrionCop.BlnFacturando Then
            Dim ldtmFechaMin As Date = GCDTMFECHANULA
            Dim ldtmFechaMax As Date = Date.Today
            If Not ClsOrionCop.BlnProcesoEspecial Then
                ldtmFechaMin = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
                ldtmFechaMax = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
            End If
            HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax, BlnEsRequerido)
            If Not HblnEsValido Then
                If Not ClsOrionCop.BlnFacturando AndAlso
                        MobjPadre.ObjIdModoFacturacionByt.ObjValorPro = EnuModoFacturacionDef.EnuManual Then
                    HstrMens = "La Fecha de Factura está por fuera del Período Actual!"
                    SNotifiqueDatInv()
                End If
            End If
        Else
            HblnEsValido = (HobjValorNew = HobjValorOriginal)
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
Friend Class ClsFechaGraciaDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaPagoSinMora"
    Private ReadOnly MobjPadre As ClsFactura = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Fecha Pago sin Mora"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando OrElse
                ClsOrionCop.BlnFacturando Then
            Dim ldtmFechaMin As Date = GCDTMFECHANULA
            Dim ldtmFechaMax As Date = Date.Today.AddYears(2)
            HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax, BlnEsRequerido)
            HstrMens = String.Empty
            If HblnEsValido Then
                If HobjValorNew <> GCDTMFECHANULA Then
                    HblnEsValido = (HobjValorNew >= MobjPadre.ObjFechaVencimientoDtm.ObjValorPro)
                    If Not HblnEsValido Then
                        HstrMens = "La Fecha del Périodo de Gracia es anterior a la Fecha de vencimiento!"
                    End If
                End If
            Else
                If Not ClsOrionCop.BlnFacturando AndAlso
                        MobjPadre.ObjIdModoFacturacionByt.ObjValorPro = EnuModoFacturacionDef.EnuManual Then
                    HstrMens = "La Fecha está por fuera del Período Actual!"
                End If
            End If
            If Not String.IsNullOrEmpty(HstrMens) Then
                SNotifiqueDatInv()
            End If
        Else
            HblnEsValido = (HobjValorNew = HobjValorOriginal)
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
Friend Class ClsFechaVencimientoDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaVencimiento"
    Private ReadOnly MobjPadre As ClsFactura = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "FechaVencimiento"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando OrElse
                ClsOrionCop.BlnFacturando Then
            Dim ldtmFechaMin As Date = MobjPadre.ObjFechaFacturaDtm.ObjValorPro
            Dim ldtmFechaMax As Date = ldtmFechaMin.AddYears(3)
            HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax, BlnEsRequerido)
            If Not HblnEsValido Then
                If Not ClsOrionCop.BlnFacturando AndAlso
                        MobjPadre.ObjIdModoFacturacionByt.ObjValorPro = EnuModoFacturacionDef.EnuManual Then
                    HstrMens = "La Fecha de Vencimiento está por fuera del Rango permitido!"
                    SNotifiqueDatInv()
                End If
            End If
        Else
            HblnEsValido = (HobjValorNew = HobjValorOriginal)
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
    Private Sub ClsFechaVencimientoDtm_evnPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            MobjPadre.ObjFechaGraciaDtm.SValide()
        End If
    End Sub
End Class
Friend Class ClsIdCliente_FactDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTerceroCliente"
    Private ReadOnly MobjPadre As ClsFactura = Nothing
    Private MobjCliente As ClsCliente = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdTerceroCliente_Factura"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjValorIng = HobjValorNew
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCDBLMINTERC, GCDBLMAXTERC,
                BlnEsRequerido)
        HstrMens = String.Empty
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                Dim lobjLlavePrincipal() = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
                MobjCliente = New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
                MobjCliente.SAbra(lobjLlavePrincipal)
                HblnEsValido = MobjCliente.BlnExiste
                If Not MobjCliente.BlnExiste Then
                    HstrMens = "El Cliente ingresado no es existe!"
                End If
                MobjPadre.ObjClienteFactura = MobjCliente
            End If
        ElseIf Not String.IsNullOrEmpty(lobjValorIng.ToString) Then
            HstrMens = "La Id. del Cliente ingresada, '" & lobjValorIng.ToString & ",  no es válida!"
        Else
            HobjValorNew = 0
        End If
        If Not String.IsNullOrEmpty(HstrMens) Then
            SNotifiqueDatInv()
        End If
    End Sub
    'Friend ReadOnly Property StrNombreCliente As String
    '    Get
    '        If Not IsNothing(MobjCliente) AndAlso MobjCliente.BlnExiste Then
    '            Return MobjCliente.ObjNombreCompletoStr.ObjValorPro
    '        Else
    '            Return ""
    '        End If
    '    End Get
    'End Property
    Private Sub ClsIdCliente_FactDbl_EvnPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If HblnEsValido Then
            MobjPadre.ObjIdPredioAgrupador_FacStr.SValide()
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
Friend Class ClsIdFacturaEnt
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsFactura = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "IdFactura"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdFactura"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 3
    End Sub
    Public Overrides Sub SValide()
        Dim lobjValorIng As Object = HobjValorNew
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue, BlnEsRequerido,
                       EnuTipoValor)
        If HblnEsValido Then
            If Not (BlnLeyendoOrigen OrElse ClsOrionCop.BlnFacturando) Then
                If ObjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                    If MobjPadre.ObjPrefijo_FactStr.BlnEsValido Then
                        Dim lstrPref = MobjPadre.ObjPrefijo_FactStr.ToString()
                        Dim lobjValorLlave() = {GshrIdCarpeta, GshrIdCentroUtil, lstrPref,
                                HobjValorNew}
                        If ObjPadre.FblnExisteLlave(lobjValorLlave) Then
                            HstrMens = "La Id. de la Factura ingresada, '" &
                                HobjValorNew.ToString & "', ya existe!"
                            HblnEsValido = False
                        End If
                    Else
                        HblnEsValido = False
                    End If
                ElseIf ObjPadre.EnuEstadoActualizacion =
                            EnuEstadoObjetoDef.enuModificando Then
                    HblnEsValido = HobjValorOriginal = HobjValorNew
                    If Not HblnEsValido Then
                        HstrMens = "No es permitido cambiar la identidad a objeto alguno!"
                    End If
                End If
            End If
        Else
            If ObjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando AndAlso
                    Not String.IsNullOrEmpty(HobjValorNew) Then
                If MobjPadre.ObjIdModoFacturacionByt.ObjValorPro IsNot Nothing AndAlso
                        (MobjPadre.ObjIdModoFacturacionByt.ObjValorPro =
                        EnuModoFacturacionDef.EnuImportada OrElse
                        MobjPadre.ObjIdModoFacturacionByt.ObjValorPro =
                        EnuModoFacturacionDef.EnuContingencia) Then
                    HstrMens = "El valor ingresado, '" & HobjValorNew.ToString &
                        "', no es valido!"
                End If
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
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdFormaPagoByt
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsFactura = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "IdFormaPago"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Id Forma Pago"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuFormaPago.EnuContado,
                    EnuFormaPago.EnuCredito, BlnEsRequerido)
        If HblnEsValido AndAlso MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            HblnEsRequerido = GobjParametros.BlnEFacAutorizado
            HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuFormaPago.EnuContado,
                    EnuFormaPago.EnuCredito, BlnEsRequerido)
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            MobjPadre.ObjIdMedioPagoByt.SValide()
        End If
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return CType(HobjValorPro, Integer).ToString
        End If
    End Function
End Class
Friend Class ClsIdInformeCont_FacEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdInformeCont"
    Private ReadOnly MobjPadre As ClsFactura = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Informe Contingencia"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        HblnEsRequerido = (MobjPadre.ObjIdModoFacturacionByt.ObjValorPro =
                EnuModoFacturacionDef.EnuContingencia)
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue, BlnEsRequerido,
                EnuTipoValor.enuInteger)
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            If HblnEsValido AndAlso GobjParametros.BlnEFacAutorizado Then
                Dim lobjValorLlave As Object = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
                Dim lobjInfCont As New ClsInformeCont(EnuModoInstanciaObjDef.enuUnico)
                lobjInfCont.SAbra(lobjValorLlave)
                If lobjInfCont.BlnExiste Then
                    HblnEsValido = (MobjPadre.ObjPrefijo_FactStr.ObjValorPro =
                            lobjInfCont.ObjPrefFactContStr.ObjValorPro) AndAlso
                            (MobjPadre.ObjIdFacturaEnt.ObjValorPro >=
                            lobjInfCont.ObjIdFactContIniEnt.ObjValorPro) AndAlso
                            (MobjPadre.ObjIdFacturaEnt.ObjValorPro <=
                            lobjInfCont.ObjIdFactContFinEnt.ObjValorPro)
                    If Not HblnEsValido Then
                        HstrMens = "La Factura actual no corresponde al Informe de Contingencia reportado!"
                        SNotifiqueDatInv()
                    End If
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
            Return CType(HobjValorPro, Integer).ToString
        End If
    End Function
End Class
Friend Class ClsIdMedioPagoByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdMedioPagoAcor"
    Private ReadOnly MobjPadre As ClsFactura = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Id Medio Pago"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuTipoMedioPagoDef.EnuEfectivo,
                    EnuTipoMedioPagoDef.EnuTransferencia, BlnEsRequerido)
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            HblnEsRequerido = GobjParametros.BlnEFacAutorizado AndAlso
                    MobjPadre.ObjIdFormaPagoByt.BlnEsValido AndAlso
                    MobjPadre.ObjIdFormaPagoByt.ObjValorPro = EnuFormaPago.EnuContado
            HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuTipoMedioPagoDef.EnuEfectivo,
                            EnuTipoMedioPagoDef.EnuTransferencia, BlnEsRequerido)
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
            Return CType(HobjValorPro, Integer).ToString
        End If
    End Function
End Class
Friend Class ClsIdModoFacturacionByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdModoFacturacion"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdModoFacturacion"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuModoFacturacionDef.EnuManual,
                EnuModoFacturacionDef.EnuContingencia, BlnEsRequerido)
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
            Return ClsOrionCop.FstrNombreModoFacturacion(HobjValorPro)
        End If
    End Function
End Class
Friend Class ClsIdPredioAgrupador_FacStr
    'Herencia
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsFactura = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "IdPredioAgrupador"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdPredioAgrupador_Factura"
        HenuTipoValor = EnuTipoValor.enuString
        HshrLongitud = 20
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud,
                BlnEsRequerido)
        If HblnEsValido AndAlso Not String.IsNullOrEmpty(HobjValorNew) Then
            Dim lobjLlavePrincipal() As Object = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                HstrMens = String.Empty
                Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
                lobjPredio.SAbra(lobjLlavePrincipal)
                HblnEsValido = lobjPredio.BlnExiste
                If HblnEsValido AndAlso Not ClsOrionCop.BlnProcesoEspecial Then
                    MobjPadre.ObjReferenciaPago_FacStr.ObjValorPro = String.Empty
                    HblnEsValido = lobjPredio.ObjIdPredioStr.ObjValorPro =
                            lobjPredio.ObjIdPredioAgrupadorStr.ObjValorPro
                    If HblnEsValido Then
                        MobjPadre.ObjReferenciaPago_FacStr.ObjValorPro =
                                lobjPredio.ObjReferenciaPagoStr.ObjValorPro
                    Else
                        HstrMens = "El Predio ingresado no es un PredioAgrupador!"
                        SNotifiqueDatInv()
                    End If
                End If
            Else
                If Not GblnActualizandoApp Then
                    HblnEsValido = (HobjValorNew = HobjValorOriginal)
                End If
            End If
        End If
        If HblnEsValido AndAlso MobjPadre.ObjIdCliente_FactDbl.BlnEsValido Then
            SLevanteEveNot("", 0, EnuSeveridadNot.EnuOk)
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(ObjValorPro) Then
            Return HobjValorPro.ToString
        Else
            Return ""
        End If
    End Function
End Class
Friend Class ClsNumeroResolAutoStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "NumeroResolAuto"
    Private ReadOnly MobjPadre As ClsFactura = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Numero Resolucion Autoriza"
        HshrLongitud = 20
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            HblnEsRequerido = (GobjParametros.BlnEFacAutorizado)
        End If
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
Friend Class ClsPieFacturaDos_FactStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PieFacturaDos"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "PieFacturaDos"
        HshrLongitud = 230
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
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
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsPieFacturaUno_FactStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PieFacturaUno"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "PieFacturaUno"
        HshrLongitud = 230
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
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
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsPrefijo_FactStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Prefijo"
    Private ReadOnly MobjPadre As ClsFactura = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "PrefijoFactura"
        HshrLongitud = 5
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsLlave = True
        HbytPosicionLlave = 2
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                If MobjPadre.ObjEsPreFacturaBln.ObjValorPro Then
                    HblnEsValido = (HobjValorNew = GCSTRPREFPREFACTURA)
                ElseIf MobjPadre.ObjIdModoFacturacionByt.ObjValorPro =
                        EnuModoFacturacionDef.EnuContingencia Then
                    HblnEsValido = (HobjValorNew =
                            GobjParametros.ObjPrefijoFactContStr.ObjValorPro)
                Else
                    If Not ClsOrionCop.BlnProcesoEspecial Then
                        HblnEsValido = (HobjValorNew =
                            GobjParametros.FstrPrefijoDoc(EnuTipoDocOri.EnuFactura))
                    End If
                End If
            ElseIf MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                If ClsOrionCop.BlnFacturando Then
                    HblnEsValido = (HobjValorNew =
                            GobjParametros.FstrPrefijoDoc(EnuTipoDocOri.EnuFactura))
                Else
                    HblnEsValido = (HobjValorNew = HobjValorOriginal)
                End If
            Else
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            End If
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando AndAlso
                Not MobjPadre.BlnVaciandoObjeto Then
            MobjPadre.ObjIdFacturaEnt.SValide()
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
Friend Class ClsReferenciaPago_FacStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "ReferenciaPago"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Referencia  Pago Factura"
        HshrLongitud = 8
        HenuTipoValor = EnuTipoValor.EnuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud,
                    BlnEsRequerido)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(ObjValorPro) Then
            If String.IsNullOrEmpty(ObjValorPro.ToString()) Then
                Return "Sin Ref. de Pago"
            Else
                Return HobjValorPro.ToString
            End If
        Else
            Return "SIN R.P."
        End If
    End Function
End Class
Friend Class ClsValor_FactDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "ValorFactura"
    Private ReadOnly MobjPadre As ClsFactura = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Valor"
        HenuTipoValor = EnuTipoValor.EnuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Decimal.MaxValue,
                        BlnEsRequerido, HenuTipoValor)
        If Not HblnEsValido Then
            If HobjValorNew = 0 Then
                HblnEsValido = MobjPadre.ObjAnuladoBln.ObjValorPro
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
            Return Format(HobjValorPro, "c")
        End If
    End Function
End Class
#End Region
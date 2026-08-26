Imports System.Text
Friend Class ClsItemFactura
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriItemsFactura"
    ' Variables de modulo
    Private MobjMiFactura As ClsFactura = Nothing
    Private McolNovedades As Collection = Nothing
    Private MdtbNovedades As DataTable = Nothing
    Private MblnCalculoDeudas As Boolean = False
    Private MdecDeudaSer As Decimal = 0
    Private MdecDeudaInt As Decimal = 0
    Private MdecDeudaIvaSer As Decimal = 0
    Private MdecDeudaIvaInt As Decimal = 0
    Private MdecCrAplicadoCap As Decimal = 0
    Private MdecCrAplicadoInt As Decimal = 0
    Private MdecIvaSerAlGasto As Decimal = 0
    Private MdecIvaIntAlGasto As Decimal = 0
    Private MobjServicio As ClsServicio = Nothing
    Private MobjSector As ClsSector = Nothing
#End Region

#Region "Constructores"
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Administrador al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwObjeto">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Public Sub New(aobjPadre As ClsCBObjetoPan, adrwObjeto As DataRow)
        HobjPadre = aobjPadre
        If TypeOf aobjPadre Is ClsFactura Then
            MobjMiFactura = aobjPadre
        Else
            HblnEsAnulable = False
            HblnEsCreable = False
        End If
        HenuTipoObjeto = EnuModoInstanciaObjDef.EnuDeColeccion
        '
        HblnEsSuprimible = False
        DrwRegistroActual = adrwObjeto
        DtbTablaColeccion = DrwRegistroActual.Table
        HenuTipoPermiso = EnuPermisosDef.EnuHeredado
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
            Return EnuIdClasesPanDef.EnuItemFactura
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Item Factura"
        End Get
    End Property
#End Region

#Region "Propiedades Prop"
    Friend ReadOnly Property ObjCreditos_ItemFactDec As New ClsCreditos_ItemFactDec(Me)
    Friend ReadOnly Property ObjDebitos_ItemFactDec As New ClsDebitos_ItemFactDec(Me)
    Friend ReadOnly Property ObjDetalle_ItemFactStr As New ClsDetalle_ItemFactStr(Me)
    Friend ReadOnly Property ObjEsExcluidoIva_ItemFactBln As New ClsEsExcluidoIva_ItemFactBln(Me)
    Friend ReadOnly Property ObjEsPrefactura_ItemFactBln As New ClsEsPreFacturaBln(Me)
    Friend ReadOnly Property ObjFechaCancelacion_ItemFactDtm As New ClsFechaCancelacion_ItemFactDtm(Me)
    Friend ReadOnly Property ObjFechaCausoIntMora_Dtm As New ClsFechaCausoIntMora_Dtm(Me)
    Friend ReadOnly Property ObjFechaGraciaIFDtm As New ClsFechaGraciaIFDtm(Me)
    Friend ReadOnly Property ObjFechaVencimientoIFDtm As New ClsFechaVencimientoIFDtm(Me)
    Friend ReadOnly Property ObjIdAno_ServicioItemFactShr As New ClsIdAno_ServicioItemFactShr(Me)
    Friend ReadOnly Property ObjIdCarpeta_ItemFactShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_ItemFactShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdFactura_ItemFactEnt As New ClsIdFactura_ItemFactEnt(Me)
    Friend ReadOnly Property ObjIdItemFacturaShr As New ClsIdItemFacturaShr(Me)
    Friend ReadOnly Property ObjIdPredio_ItemFactStr As New ClsIdPredio_ItemFactStr(Me)
    Friend ReadOnly Property ObjIdServicio_ItemFactShr As New ClsIdServicio_ItemFactShr(Me)
    Friend ReadOnly Property ObjPeriodo_ItemFactStr As New ClsPeriodo_ItemFactStr(Me)
    Friend ReadOnly Property ObjPrefijo_ItemFactStr As New ClsPrefijo_ItemFactStr(Me)
    Friend ReadOnly Property ObjTarifaIva_ItemFactDbl As New ClsTarifaIva_ItemFactDbl(Me)
    Friend ReadOnly Property ObjValor_ItemFactDec As New ClsValor_ItemFactDec(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjAnuladoBln)
                HcolPropiedades.Add(ObjCreditos_ItemFactDec)
                HcolPropiedades.Add(ObjDebitos_ItemFactDec)
                HcolPropiedades.Add(ObjDetalle_ItemFactStr)
                HcolPropiedades.Add(ObjEsExcluidoIva_ItemFactBln)
                HcolPropiedades.Add(ObjEsPrefactura_ItemFactBln)
                HcolPropiedades.Add(ObjFechaCancelacion_ItemFactDtm)
                HcolPropiedades.Add(ObjFechaCausoIntMora_Dtm)
                HcolPropiedades.Add(ObjFechaGraciaIFDtm)
                HcolPropiedades.Add(ObjFechaVencimientoIFDtm)
                HcolPropiedades.Add(ObjIdAno_ServicioItemFactShr)
                HcolPropiedades.Add(ObjIdCarpeta_ItemFactShr)
                HcolPropiedades.Add(ObjIdCentroUtil_ItemFactShr)
                HcolPropiedades.Add(ObjIdFactura_ItemFactEnt)
                HcolPropiedades.Add(ObjIdItemFacturaShr)
                HcolPropiedades.Add(ObjIdPredio_ItemFactStr)
                HcolPropiedades.Add(ObjIdServicio_ItemFactShr)
                HcolPropiedades.Add(ObjPeriodo_ItemFactStr)
                HcolPropiedades.Add(ObjPrefijo_ItemFactStr)
                HcolPropiedades.Add(ObjTarifaIva_ItemFactDbl)
                HcolPropiedades.Add(ObjValor_ItemFactDec)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region

#Region "Otras propiedades"
    ''' <summary>
    ''' Devuelve uns string compuesto por el prefijo de la factura y el id de la factura separados por un
    ''' guion. Si no existe el prefijo devuelve solo el id de la factura
    ''' </summary>
    ''' <value></value>
    Friend ReadOnly Property StrNumeroFactura As String
        Get
            Dim lstrNroFactura As String = ClsPanorama.FstrNumeroDcto(ObjPrefijo_ItemFactStr.ObjValorPro,
                    ObjIdFactura_ItemFactEnt.ObjValorPro)
            Return lstrNroFactura
        End Get
    End Property

    Friend ReadOnly Property ObjServicio As ClsServicio
        Get
            If ObjIdAno_ServicioItemFactShr.BlnEsValido AndAlso ObjIdServicio_ItemFactShr.BlnEsValido Then
                Dim lshrIdAno As Short = ObjIdAno_ServicioItemFactShr.ObjValorPro
                Dim lshrIdServicio As Short = ObjIdServicio_ItemFactShr.ObjValorPro
                Dim lstrKey = lshrIdAno.ToString & "," & lshrIdServicio.ToString
                If lshrIdAno = 0 Then
                    MobjServicio = GobjParametros.ColServiciosPer(lstrKey)
                Else
                    Dim lobjAno As ClsAno = GobjParametros.ColAnos(lshrIdAno.ToString)
                    MobjServicio = lobjAno.ColServiciosAno(lstrKey)
                End If
            End If
            Return MobjServicio
        End Get
    End Property

    Friend ReadOnly Property ShrDiasGracia As Short
        Get
            Return ObjServicio.ObjDiasGraciaShr.ObjValorPro
        End Get
    End Property

    Friend ReadOnly Property ObjSector As ClsSector
        Get
            If IsNothing(MobjSector) AndAlso ObjIdPredio_ItemFactStr.ToString.Length > 0 Then
                Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.EnuUnico)
                Dim lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, ObjIdPredio_ItemFactStr.ObjValorPro}
                lobjPredio.SAbra(lobjValorLlave)
                MobjSector = lobjPredio.ObjSector
            End If
            Return MobjSector
        End Get
    End Property

    Friend ReadOnly Property StrServicio As String
        Get
            Dim lstrServ As String
            If ObjIdAno_ServicioItemFactShr.ObjValorPro > 0 Then
                lstrServ = "0"
            Else
                lstrServ = ObjIdServicio_ItemFactShr.ToString
            End If
            Return lstrServ
        End Get
    End Property

    Friend ReadOnly Property ObjMiFactura As ClsFactura
        Get
            If BlnExiste AndAlso IsNothing(MobjMiFactura) Then
                Dim lobjFactura As New ClsFactura()
                lobjFactura.SAbra({GshrIdCarpeta, GshrIdCentroUtil,
                        ObjPrefijo_ItemFactStr.ObjValorPro, ObjIdFactura_ItemFactEnt.ObjValorPro})
                MobjMiFactura = lobjFactura
            End If
            Return MobjMiFactura
        End Get
    End Property

    Friend ReadOnly Property BlnDeudaEsCuotaAdmin As Boolean
        Get
            Dim lblnEs As Boolean = False
            If ObjDebitos_ItemFactDec.ObjValorPro > ObjCreditos_ItemFactDec.ObjValorPro Then
                lblnEs = If(ObjIdAno_ServicioItemFactShr.ObjValorPro > 0, True, False)
            End If
            Return lblnEs
        End Get
    End Property
#End Region
#End Region

#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SVacie()
        MyBase.SVacie()
        MdtbNovedades = Nothing
        McolNovedades = Nothing
        MobjServicio = Nothing
        MobjSector = Nothing
        MblnCalculoDeudas = False
        MdecDeudaSer = 0
        MdecDeudaInt = 0
        MdecDeudaIvaSer = 0
        MdecDeudaIvaInt = 0
        MdecCrAplicadoCap = 0
        MdecCrAplicadoInt = 0
        MdecIvaSerAlGasto = 0
        MdecIvaIntAlGasto = 0
    End Sub
    Protected Overrides Sub SInicialiceObj()
        ObjIdCarpeta_ItemFactShr.ObjValorPro = GshrIdCarpeta
        ObjIdCentroUtil_ItemFactShr.ObjValorPro = GshrIdCentroUtil
        ObjIdFactura_ItemFactEnt.ObjValorPro = 0
        ObjAnuladoBln.ObjValorPro = False
        ObjCreditos_ItemFactDec.ObjValorPro = 0
        ObjFechaCancelacion_ItemFactDtm.ObjValorPro = GCDTMFECHANULA
        ObjPeriodo_ItemFactStr.ObjValorPro = GobjParametros.ObjAnoActual.StrIdPeriodoActual
        ObjFechaCausoIntMora_Dtm.ObjValorPro = GCDTMFECHANULA
    End Sub
    Protected Overrides Sub SLeaValores(ablnLeyendoOrigen As Boolean)
        MyBase.SLeaValores(ablnLeyendoOrigen)
        If BlnExiste AndAlso IsNothing(MobjMiFactura) Then
            Dim lobjFactura As New ClsFactura()
            lobjFactura.SAbra({GshrIdCarpeta, GshrIdCentroUtil, ObjPrefijo_ItemFactStr.ObjValorPro,
                    ObjIdFactura_ItemFactEnt.ObjValorPro})
            MobjMiFactura = lobjFactura
        End If
    End Sub
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        GobjPanDat.SControleProcesoObj(True)
        Try
            If EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                ' mcolNovedades puede ser Nothing en el proceso de causar intereses de mora porque
                ' la novedad es generada en la Nota Db
                If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                    SGenereNovedades()
                    If Not ObjServicio.ObjEsServicioIdBln.ObjValorPro Then
                        SAsigneFacturaNovedad()
                    End If
                End If
                If Not IsNothing(McolNovedades) Then
                    ClsPanorama.SActualiceCol(McolNovedades)
                End If
                ObjDebitos_ItemFactDec.SValide()
                ObjValor_ItemFactDec.SValide()
                MyBase.SActualice(ablnExigeRequeridos)
                SVacie()
            End If
        Catch ex As PanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        Finally
            GobjPanDat.SControleProcesoObj(False)
        End Try
    End Sub
    Protected Overrides Function SAnuleEnObj() As Boolean
        EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
        ObjAnuladoBln.ObjValorPro = True
        Return True
    End Function
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjIdItemFacturaShr.ToString
        End Get
    End Property
#End Region

#Region "Procedimientos del objeto"
    Friend Sub SModifiqueADefinitiva(astrPrefijo As String)
        For Each lobjNovedad As ClsNovedad In ColNovedades
            lobjNovedad.SModifiqueADefinitiva(astrPrefijo)
        Next
        EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
        ObjEsPrefactura_ItemFactBln.ObjValorPro = False
        ObjPrefijo_ItemFactStr.ObjValorPro = astrPrefijo
    End Sub
    Friend Function FblnRetencionAplicada(aenuTipoDscto As EnuTipoDescuento) As Boolean
        Dim lblnRetApli As Boolean, ldecVlrRet As Decimal
        Dim lenuTipoNov As EnuTipoNov
        For Each lobjNov As ClsNovedad In ColNovedades
            lenuTipoNov = lobjNov.ObjIdTipoNovedadByt.ObjValorPro
            Select Case aenuTipoDscto
                Case EnuTipoDescuento.EnuReteFuente
                    If lenuTipoNov = EnuTipoNov.EnuCrRetFte Then
                        ldecVlrRet += lobjNov.ObjValor_NovDec.ObjValorPro
                    ElseIf lenuTipoNov = EnuTipoNov.EnuRCrRetFte Then
                        ldecVlrRet -= lobjNov.ObjValor_NovDec.ObjValorPro
                    End If
                Case EnuTipoDescuento.EnuReteIca
                    If lenuTipoNov = EnuTipoNov.EnuCrRetIca Then
                        ldecVlrRet += lobjNov.ObjValor_NovDec.ObjValorPro
                    ElseIf lenuTipoNov = EnuTipoNov.EnuRCrRetIca Then
                        ldecVlrRet -= lobjNov.ObjValor_NovDec.ObjValorPro
                    End If
                Case EnuTipoDescuento.EnuReteIva
                    If lenuTipoNov = EnuTipoNov.EnuCrRetIva Then
                        ldecVlrRet += lobjNov.ObjValor_NovDec.ObjValorPro
                    ElseIf lenuTipoNov = EnuTipoNov.EnuRCrRetIva Then
                        ldecVlrRet -= lobjNov.ObjValor_NovDec.ObjValorPro
                    End If
            End Select
        Next
        lblnRetApli = (ldecVlrRet > 0)
        Return lblnRetApli
    End Function
    ''' <summary>
    ''' Devuelve el valor recibido para ser aplicado al capital y a los intereses de la deuda
    ''' </summary>
    Friend Function FdecPagoTotalAplicado() As Decimal
        Dim ldecVlrPago = 0D
        Dim lenuTipoNov As EnuTipoNov
        For Each lobjNov As ClsNovedad In ColNovedades
            lenuTipoNov = lobjNov.ObjIdTipoNovedadByt.ObjValorPro
            Select Case lenuTipoNov
                Case EnuTipoNov.enuCrPagoCap, EnuTipoNov.enuCrAnApCap
                    ldecVlrPago += lobjNov.ObjValor_NovDec.ObjValorPro
                Case EnuTipoNov.enuCrPagoInt, EnuTipoNov.enuCrAnApInt
                    ldecVlrPago += lobjNov.ObjValor_NovDec.ObjValorPro
                Case EnuTipoNov.enuRCrPagoCap, EnuTipoNov.enuRCrAnApCap
                    ldecVlrPago -= lobjNov.ObjValor_NovDec.ObjValorPro
                Case EnuTipoNov.enuRCrPagoInt, EnuTipoNov.enuRCrAnApInt
                    ldecVlrPago -= lobjNov.ObjValor_NovDec.ObjValorPro
            End Select
        Next
        Return ldecVlrPago
    End Function
#End Region

#Region "Aplica Creditos (Pagos, Descuentos y retenciones)"
    Friend Sub SApliqueCredito(aenuTipoNovedad As EnuTipoNov, adtmFechaCredito As Date,
            astrIdCtaDebito As String, ashrIdItemDocOrigen As Short,
            astrPrefijoDocOrigen As String, aentIdDocOrigen As Integer,
            aenuTipoDocOrigen As EnuTipoDocOri, adecValor As Decimal, adecBase As Decimal,
            adblFactor As Double)
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
            EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
        End If
        If astrIdCtaDebito = GCSTRCUENTADSCTOCAP Then
            If aenuTipoNovedad = EnuTipoNov.EnuRDbIva Then
                astrIdCtaDebito = ObjServicio.ObjCodigoCuentaIvaStr.ObjValorPro
            Else
                astrIdCtaDebito = ObjServicio.ObjCodigoCuentaDevStr.ObjValorPro
            End If
        ElseIf astrIdCtaDebito = GCSTRCUENTADSCTOINT Then
            astrIdCtaDebito = ObjServicio.ObjCodigoCuentaMoraStr.ObjValorPro
        End If
        ' Descuento al servicio
        Dim lstrIdCtaCr As String, lshrIdSer As Short, lshrIdAno As Short
        Select Case aenuTipoNovedad
            Case EnuTipoNov.EnuCrPagoInt, EnuTipoNov.EnuCrAnApInt, EnuTipoNov.EnuCrDctoInt
                lstrIdCtaCr = GobjParametros.ObjIdCtaIntMoraDbStr.ObjValorPro
                lshrIdSer = GCSHRIDMORA
                lshrIdAno = 0
            Case Else
                lstrIdCtaCr = ObjServicio.ObjCodigoCuentaDbStr.ObjValorPro
                lshrIdSer = ObjServicio.ObjIdServicioShr.ObjValorPro
                lshrIdAno = ObjServicio.ObjIdAno_ServicioShr.ObjValorPro
        End Select
        SGenereNovCredito(adtmFechaCredito, astrIdCtaDebito, lstrIdCtaCr, aentIdDocOrigen,
                ashrIdItemDocOrigen, aenuTipoDocOrigen, aenuTipoNovedad,
                astrPrefijoDocOrigen, adecValor, adecBase, adblFactor, lshrIdAno, lshrIdSer)
        ObjCreditos_ItemFactDec.ObjValorPro += adecValor
        SRegistreCr(adecValor, aenuTipoNovedad)
        SVerifiqueCancelacion(adtmFechaCredito)
    End Sub
    Private Sub SGenereNovCredito(adtmFechaCredito As Date, astrIdCtaDebito As String,
                astrIdCtaCredito As String, aentIdDocOrigen As Integer,
                ashrIdItemDocOrigen As Short, aenuTipoDocOrigen As EnuTipoDocOri,
                aenuTipoNovedad As EnuTipoNov, astrPrefijoDocOrigen As String,
                adecValor As Decimal, adecBase As Decimal, adblFactor As Double,
                ashrIdAno As Short, ashrIdServicio As Short)
        Dim lobjNovedad As ClsNovedad = FobjNuevaNovedad()
        With lobjNovedad
            .ObjBaseDec.ObjValorPro = adecBase
            .ObjFactorDbl.ObjValorPro = adblFactor
            .ObjFechaNovedadDtm.ObjValorPro = adtmFechaCredito
            .ObjIdCuentaDb_NovStr.ObjValorPro = astrIdCtaDebito
            .ObjIdCuentaCr_NovStr.ObjValorPro = astrIdCtaCredito
            .ObjIdDocOrigenEnt.ObjValorPro = aentIdDocOrigen
            .ObjIdPredioAgrupador_NovStr.ObjValorPro = MobjMiFactura.ObjIdPredioAgrupador_FacStr.ObjValorPro
            .ObjPrefijoFact_NovStr.ObjValorPro = MobjMiFactura.ObjPrefijo_FactStr.ObjValorPro
            .ObjIdFactura_NovEnt.ObjValorPro = ObjIdFactura_ItemFactEnt.ObjValorPro
            .ObjIdItemDocOrigen_NovShr.ObjValorPro = ashrIdItemDocOrigen
            .ObjIdTipoDocOrigenByt.ObjValorPro = aenuTipoDocOrigen
            .ObjIdTipoNovedadByt.ObjValorPro = aenuTipoNovedad
            .ObjPrefijoDocOrigen_NovStr.ObjValorPro = astrPrefijoDocOrigen
            .ObjValor_NovDec.ObjValorPro = adecValor
            .ObjIdAno_NovShr.ObjValorPro = ashrIdAno
            .ObjIdServicio_NovShr.ObjValorPro = ashrIdServicio
        End With
        ColNovedades.Add(lobjNovedad)
    End Sub
    Private Sub SVerifiqueCancelacion(adtmFechaCredito As Date)
        If ObjCreditos_ItemFactDec.ObjValorPro > ObjDebitos_ItemFactDec.ObjValorPro Then
            Dim lstrMens = "Creditos mayores a los debitos en el Item " & ObjIdItemFacturaShr.ToString &
                    " de la Factura " & ObjIdFactura_ItemFactEnt.ToString & "!"
            Throw New ErrorInesperadoPanLException(lstrMens)
        End If
        If ObjCreditos_ItemFactDec.ObjValorPro = ObjDebitos_ItemFactDec.ObjValorPro Then
            ObjFechaCancelacion_ItemFactDtm.ObjValorPro = adtmFechaCredito
        End If
    End Sub
    Private Sub SRegistreCr(adecValor As Decimal, aenuTipoNov As EnuTipoNov)
        Select Case aenuTipoNov
            Case EnuTipoNov.EnuCrAnApCap, EnuTipoNov.EnuCrDctoCap, EnuTipoNov.EnuCrPagoCap,
                    EnuTipoNov.EnuCrRetCre, EnuTipoNov.EnuCrRetFte, EnuTipoNov.EnuCrRetIca,
                    EnuTipoNov.EnuCrRetIva
                MdecCrAplicadoCap += adecValor
            Case EnuTipoNov.EnuCrAnApInt, EnuTipoNov.EnuCrDctoInt, EnuTipoNov.EnuCrPagoInt
                MdecCrAplicadoInt += adecValor
        End Select
    End Sub
#End Region

#Region "Manejo Novedades"
    Friend ReadOnly Property ColNovedades As Collection
        Get
            If IsNothing(McolNovedades) OrElse McolNovedades.Count = 0 Then
                McolNovedades = New Collection
                SCargueDtbNovedades()
                If Not IsNothing(MdtbNovedades) AndAlso MdtbNovedades.Rows.Count > 0 Then
                    Dim ldrwNovedades() As DataRow = MdtbNovedades.Select
                    For Each ldrwNovedad As DataRow In ldrwNovedades
                        Dim lobjNovedad As New ClsNovedad(Me, ldrwNovedad)
                        lobjNovedad.SLeaValores(True)
                        McolNovedades.Add(lobjNovedad, lobjNovedad.ObjIdNovedadShr.ToString)
                    Next
                End If
            End If
            Return McolNovedades
        End Get
    End Property
    Private Sub SCargueDtbNovedades()
        If IsNothing(MdtbNovedades) Then
            Dim lstrIdFactura = ObjIdFactura_ItemFactEnt.ToString
            Dim lstrPrefijoFac = ObjPrefijo_ItemFactStr.ToString
            Dim lstrIndice = {{ClsIdNovedadShr.SstrNombreCampoBd, "ASC"}}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsPrefijoFact_NovStr.SstrNombreCampoBd & " = '" & lstrPrefijoFac &
                    "' AND " & ClsIdFactura_NovEnt.SstrNombreCampoBd & " = " & lstrIdFactura &
                    " AND " & ClsIdItemFacturaShr.SstrNombreCampoBd & " = " &
                    ObjIdItemFacturaShr.ObjValorPro
            Dim lstrCamposSelect() = {"*"}
            MdtbNovedades = ClsPanorama.FdtbDataTable(ClsNovedad.SstrNombreTabla, lstrCamposSelect,
                    lstrIndice, lstrFiltro)
        End If
    End Sub
    Friend Sub SGenereNovedades()
        Try
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                If IsNothing(McolNovedades) Then
                    McolNovedades = ColNovedades
                End If
                SGenereNovedad_CxC_Ing()
                If DecIvaServicio() > 0 Then
                    SGenereNovedad_CxC_Iva()
                End If
            End If
        Catch ex As ErrorInesperadoPanLException
            Throw
        Catch es As ErrorInesperadoPanDatException
            Throw
        Catch es As PanDatException
            Throw
        Catch ex As PanLException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Sub
    Private Sub SGenereNovedad_CxC_Ing()
        ' Novedad: Debito=CxC Crédito=Ing o Cta Por Pagar
        Dim lobjNovedad As ClsNovedad = FobjNuevaNovedad()
        With lobjNovedad
            .ObjBaseDec.ObjValorPro = 0
            .ObjFactorDbl.ObjValorPro = 0
            .ObjIdCuentaDb_NovStr.ObjValorPro = ObjServicio.ObjCodigoCuentaDbStr.ObjValorPro
            .ObjIdCuentaCr_NovStr.ObjValorPro = ObjServicio.ObjCodigoCuentaCrStr.ObjValorPro
            .ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuDbCap
            .ObjValor_NovDec.ObjValorPro = DecValorServicio()
            .ObjIdAno_NovShr.ObjValorPro = ObjServicio.ObjIdAno_ServicioShr.ObjValorPro
            .ObjIdServicio_NovShr.ObjValorPro = ObjServicio.ObjIdServicioShr.ObjValorPro
            If ObjServicio.ObjEsServicioIdBln.ObjValorPro Then
                Dim lobjFac As ClsFactura = ObjPadre
                .ObjIdPredioAgrupador_NovStr.ObjValorPro = lobjFac.ObjIdPredioAgrupador_FacStr.ObjValorPro
                .ObjPrefijoDocOrigen_NovStr.ObjValorPro = ObjPrefijo_ItemFactStr.ObjValorPro
                .ObjPrefijoFact_NovStr.ObjValorPro = ObjPrefijo_ItemFactStr.ObjValorPro
                .ObjIdDocOrigenEnt.ObjValorPro = ObjIdFactura_ItemFactEnt.ObjValorPro
                .ObjIdFactura_NovEnt.ObjValorPro = ObjIdFactura_ItemFactEnt.ObjValorPro
            End If
            .ObjIdTerceroCtaCr_NovDbl.ObjValorPro = ObjServicio.ObjIdTerceroCtaCrDbl.ObjValorPro
        End With
        McolNovedades.Add(lobjNovedad)
    End Sub
    Private Sub SGenereNovedad_CxC_Iva()
        ' Novedad: Debito=CxC Crédito=Iva
        Dim lobjNovedad As ClsNovedad = FobjNuevaNovedad()
        With lobjNovedad
            .ObjBaseDec.ObjValorPro = DecBaseIvaServicio()
            .ObjFactorDbl.ObjValorPro = ObjServicio.ObjTarifaIvaDbl.ObjValorPro
            .ObjIdCuentaDb_NovStr.ObjValorPro = ObjServicio.ObjCodigoCuentaDbStr.ObjValorPro
            .ObjIdCuentaCr_NovStr.ObjValorPro = ObjServicio.ObjCodigoCuentaIvaStr.ObjValorPro
            .ObjIdAno_NovShr.ObjValorPro = ObjServicio.ObjIdAno_ServicioShr.ObjValorPro
            .ObjIdServicio_NovShr.ObjValorPro = ObjServicio.ObjIdServicioShr.ObjValorPro
            .ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuDbIva
            .ObjValor_NovDec.ObjValorPro = DecIvaServicio()
        End With
        McolNovedades.Add(lobjNovedad)
    End Sub
    ''' <summary>
    ''' Crea un nueva novedad con las cuentas invertidas por cada novedad generada al ser creada la factura 
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub SAnuleNovedadesItem(aobjItemNotaCr As ClsItemNotaCr)
        Static lshrIdProxNov As Short = MobjMiFactura.FshrIdUltimaNov + 1
        Dim lcolNewNovedades As New Collection
        Dim lobjNovedad As ClsNovedad = Nothing, lobjFactura As ClsFactura = ObjPadre
        Dim lstrCtaCr = String.Empty, lstrCtaDb = String.Empty
        Dim lobjNotaCr As ClsNotaCr = aobjItemNotaCr.ObjPadre
        Dim lenuTipoNovedad As EnuTipoNov = EnuTipoNov.None
        Dim lenuTipoDscto As EnuTipoDescuento =
                aobjItemNotaCr.ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro
        Dim lshrIdItemNCR As Short = aobjItemNotaCr.ObjIdItemNotaCrShr.ObjValorPro
        Dim lblnEsRevIva As Boolean = aobjItemNotaCr.ObjEsReversionIvaBln.ObjValorPro
        If lenuTipoDscto <> EnuTipoDescuento.EnuDsctoIntMora Then
            For Each lobjNov As ClsNovedad In ColNovedades
                lenuTipoNovedad = lobjNov.ObjIdTipoNovedadByt.ObjValorPro
                If ((lenuTipoNovedad = EnuTipoNov.EnuDbCap AndAlso Not lblnEsRevIva) OrElse
                        (lenuTipoNovedad = EnuTipoNov.EnuDbIva AndAlso lblnEsRevIva)) Then
                    lstrCtaCr = lobjNov.ObjIdCuentaDb_NovStr.ObjValorPro
                    lstrCtaDb = lobjNov.ObjIdCuentaCr_NovStr.ObjValorPro
                    lobjNovedad = FobjNuevaNovedad()
                    With lobjNovedad
                        .ObjIdPredioAgrupador_NovStr.ObjValorPro =
                                lobjFactura.ObjIdPredioAgrupador_FacStr.ObjValorPro
                        .ObjIdFactura_NovEnt.ObjValorPro = ObjIdFactura_ItemFactEnt.ObjValorPro
                        .ObjPrefijoFact_NovStr.ObjValorPro = ObjPrefijo_ItemFactStr.ObjValorPro
                        .ObjFechaNovedadDtm.ObjValorPro = lobjNotaCr.ObjFecha_NotaCrDtm.ObjValorPro
                        .ObjIdTipoDocOrigenByt.ObjValorPro = EnuTipoDocOri.EnuNotaCr
                        .ObjPrefijoDocOrigen_NovStr.ObjValorPro =
                                lobjNotaCr.ObjPrefijo_NotaCrStr.ObjValorPro
                        .ObjIdDocOrigenEnt.ObjValorPro = lobjNotaCr.ObjIdNotaCrEnt.ObjValorPro
                        .ObjIdItemDocOrigen_NovShr.ObjValorPro = lshrIdItemNCR
                        .ObjBaseDec.ObjValorPro = aobjItemNotaCr.ObjBaseDscto_NotaCrDec.ObjValorPro
                        .ObjFactorDbl.ObjValorPro =
                                aobjItemNotaCr.ObjTasaDscto_ItemNotaCrDbl.ObjValorPro
                        .ObjIdCuentaDb_NovStr.ObjValorPro = lstrCtaDb
                        .ObjIdCuentaCr_NovStr.ObjValorPro = lstrCtaCr
                        .ObjIdTipoNovedadByt.ObjValorPro = ClsOrionCop.FenuTipoNovContraria(
                                lobjNov.ObjIdTipoNovedadByt.ObjValorPro)
                        .ObjValor_NovDec.ObjValorPro =
                                aobjItemNotaCr.ObjValor_ItemNotaCrDec.ObjValorPro
                        .ObjIdAno_NovShr.ObjValorPro = lobjNov.ObjIdAno_NovShr.ObjValorPro
                        .ObjIdServicio_NovShr.ObjValorPro = lobjNov.ObjIdServicio_NovShr.ObjValorPro
                    End With
                    lcolNewNovedades.Add(lobjNovedad)
                    ObjCreditos_ItemFactDec.ObjValorPro += lobjNov.ObjValor_NovDec.ObjValorPro
                    Exit For
                End If
            Next
        Else
            For Each lobjNov As ClsNovedad In ColNovedades
                lenuTipoNovedad = lobjNov.ObjIdTipoNovedadByt.ObjValorPro
                If ((lenuTipoNovedad = EnuTipoNov.EnuDbInt AndAlso Not lblnEsRevIva) OrElse
                        (lenuTipoNovedad = EnuTipoNov.EnuDbIvaInt AndAlso lblnEsRevIva)) Then
                    lstrCtaCr = lobjNov.ObjIdCuentaDb_NovStr.ObjValorPro
                    lstrCtaDb = lobjNov.ObjIdCuentaCr_NovStr.ObjValorPro
                    lobjNovedad = FobjNuevaNovedad()
                    With lobjNovedad
                        .ObjIdPredioAgrupador_NovStr.ObjValorPro =
                                lobjFactura.ObjIdPredioAgrupador_FacStr.ObjValorPro
                        .ObjIdFactura_NovEnt.ObjValorPro = ObjIdFactura_ItemFactEnt.ObjValorPro
                        .ObjPrefijoFact_NovStr.ObjValorPro = ObjPrefijo_ItemFactStr.ObjValorPro
                        .ObjFechaNovedadDtm.ObjValorPro = lobjNotaCr.ObjFecha_NotaCrDtm.ObjValorPro
                        .ObjIdTipoDocOrigenByt.ObjValorPro = EnuTipoDocOri.EnuNotaCr
                        .ObjPrefijoDocOrigen_NovStr.ObjValorPro =
                                lobjNotaCr.ObjPrefijo_NotaCrStr.ObjValorPro
                        .ObjIdDocOrigenEnt.ObjValorPro = lobjNotaCr.ObjIdNotaCrEnt.ObjValorPro
                        .ObjIdItemDocOrigen_NovShr.ObjValorPro = lshrIdItemNCR
                        .ObjBaseDec.ObjValorPro = aobjItemNotaCr.ObjBaseDscto_NotaCrDec.ObjValorPro
                        .ObjFactorDbl.ObjValorPro =
                                aobjItemNotaCr.ObjTasaDscto_ItemNotaCrDbl.ObjValorPro
                        .ObjIdCuentaDb_NovStr.ObjValorPro = lstrCtaDb
                        .ObjIdCuentaCr_NovStr.ObjValorPro = lstrCtaCr
                        .ObjIdTipoNovedadByt.ObjValorPro = ClsOrionCop.FenuTipoNovContraria(
                                lobjNov.ObjIdTipoNovedadByt.ObjValorPro)
                        .ObjValor_NovDec.ObjValorPro =
                                aobjItemNotaCr.ObjValor_ItemNotaCrDec.ObjValorPro
                        .ObjIdAno_NovShr.ObjValorPro = lobjNov.ObjIdAno_NovShr.ObjValorPro
                        .ObjIdServicio_NovShr.ObjValorPro = lobjNov.ObjIdServicio_NovShr.ObjValorPro
                    End With
                    lcolNewNovedades.Add(lobjNovedad)
                    ObjCreditos_ItemFactDec.ObjValorPro +=
                            aobjItemNotaCr.ObjValor_ItemNotaCrDec.ObjValorPro
                    Exit For
                End If
            Next
        End If
        For Each lobjNov As ClsNovedad In lcolNewNovedades
            lobjNov.ObjIdNovedadShr.ObjValorPro = lshrIdProxNov
            ColNovedades.Add(lobjNov, lshrIdProxNov.ToString)
            lshrIdProxNov += 1
        Next
        lcolNewNovedades.Clear()
        lcolNewNovedades = Nothing
    End Sub
    Friend Function FobjNuevaNovedad() As ClsNovedad
        Dim lobjNovedad As ClsNovedad = Nothing
        If EnuEstadoActualizacion <> EnuEstadoObjetoDef.EnuConsultando Then
            SCargueDtbNovedades()
            If Not IsNothing(MdtbNovedades) Then
                Dim lblnModificoPermisos = False
                Dim ldrwNovedad As DataRow = MdtbNovedades.NewRow
                lobjNovedad = New ClsNovedad(Me, ldrwNovedad)
                With lobjNovedad
                    If Not CType(.EnuPermisosObj And EnuPermisosDef.EnuCrear, Boolean) Then
                        .EnuPermisosObj += EnuPermisosDef.EnuCrear
                        lblnModificoPermisos = True
                    End If
                    .SCreeObj(Nothing)
                    .ObjAnuladoBln.ObjValorPro = False
                    .ObjEsPrefactura_NovBln.ObjValorPro = ObjEsPrefactura_ItemFactBln.ObjValorPro
                    .ObjFechaCreacionDtm.ObjValorPro = Date.Now
                    .ObjFechaNovedadDtm.ObjValorPro = MobjMiFactura.ObjFechaFacturaDtm.ObjValorPro
                    .ObjIdCarpeta_NovShr.ObjValorPro = GshrIdCarpeta
                    .ObjIdCentroUtil_NovShr.ObjValorPro = GshrIdCentroUtil
                    .ObjIdItemFact_NovShr.ObjValorPro = ObjIdItemFacturaShr.ObjValorPro
                    .ObjIdItemDocOrigen_NovShr.ObjValorPro = ObjIdItemFacturaShr.ObjValorPro
                    .ObjIdTercero_NovDbl.ObjValorPro = MobjMiFactura.ObjIdCliente_FactDbl.ObjValorPro
                    .ObjAliasCont_NovStr.ObjValorPro = MobjMiFactura.FstrAliasCon
                    .ObjIdTipoDocOrigenByt.ObjValorPro = EnuTipoDocOri.EnuFactura
                    .ObjPrefijoFact_NovStr.ObjValorPro = Me.ObjPrefijo_ItemFactStr.ObjValorPro
                    .ObjPrefijoDocOrigen_NovStr.ObjValorPro = Me.ObjPrefijo_ItemFactStr.ObjValorPro
                    .ObjIdTerceroCtaCr_NovDbl.ObjValorPro = 0
                    If lblnModificoPermisos Then
                        .EnuPermisosObj -= EnuPermisosDef.EnuCrear
                    End If
                End With
            End If
        End If
        Return lobjNovedad
    End Function
    Friend Sub SAsigneFacturaNovedad()
        If Not IsNothing(ColNovedades) AndAlso ColNovedades.Count > 0 Then
            For Each lobjNov As ClsNovedad In McolNovedades
                With lobjNov
                    .ObjIdPredioAgrupador_NovStr.ObjValorPro = MobjMiFactura.ObjIdPredioAgrupador_FacStr.ObjValorPro
                    .ObjPrefijoFact_NovStr.ObjValorPro = ObjPrefijo_ItemFactStr.ObjValorPro
                    .ObjPrefijoDocOrigen_NovStr.ObjValorPro = ObjPrefijo_ItemFactStr.ObjValorPro
                    .ObjIdFactura_NovEnt.ObjValorPro = ObjIdFactura_ItemFactEnt.ObjValorPro
                    .ObjIdDocOrigenEnt.ObjValorPro = ObjIdFactura_ItemFactEnt.ObjValorPro
                    .ObjIdItemFact_NovShr.ObjValorPro = ObjIdItemFacturaShr.ObjValorPro
                    .ObjIdItemDocOrigen_NovShr.ObjValorPro = lobjNov.ObjIdItemFact_NovShr.ObjValorPro
                    .ObjEsPrefactura_NovBln.ObjValorPro = ObjEsPrefactura_ItemFactBln.ObjValorPro
                End With
            Next
        End If
    End Sub
    Friend Sub SReverseNotaReversionRC(adecValor As Decimal, adtmFechaNovedad As Date)
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
            EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
        End If
        ObjCreditos_ItemFactDec.ObjValorPro += adecValor
        ObjFechaCausoIntMora_Dtm.ObjValorPro = adtmFechaNovedad
        If ObjDebitos_ItemFactDec.ObjValorPro = ObjCreditos_ItemFactDec.ObjValorPro Then
            ObjFechaCancelacion_ItemFactDtm.ObjValorPro = adtmFechaNovedad
        End If
    End Sub
    ''' <summary>
    ''' Crea una nueva novedad con un movimiento contrario al movimiento de la novedad pasada en el argumento
    ''' </summary>
    ''' <param name="aobjNovNovedad">Novedad a ser anulada</param>
    ''' <remarks></remarks>
    Friend Sub SReverseNovedad(aobjNovedad As ClsNovedad, adtmFechaAnu As Date,
            aenuTipoDocOrigen As EnuTipoDocOri, astrPrefijoDocOrigen As String,
            aentIdDocorigen As Integer, ashrIdItemDocOrigen As Short)
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
            EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
        End If
        Dim lobjNovedad As ClsNovedad = FobjNuevaNovedad()
        With lobjNovedad
            .ObjIdPredioAgrupador_NovStr.ObjValorPro = aobjNovedad.ObjIdPredioAgrupador_NovStr.ObjValorPro
            .ObjIdFactura_NovEnt.ObjValorPro = aobjNovedad.ObjIdFactura_NovEnt.ObjValorPro
            .ObjPrefijoFact_NovStr.ObjValorPro = aobjNovedad.ObjPrefijoFact_NovStr.ObjValorPro
            .ObjFechaNovedadDtm.ObjValorPro = adtmFechaAnu
            .ObjIdTipoDocOrigenByt.ObjValorPro = aenuTipoDocOrigen
            .ObjPrefijoDocOrigen_NovStr.ObjValorPro = astrPrefijoDocOrigen
            .ObjIdDocOrigenEnt.ObjValorPro = aentIdDocorigen
            .ObjIdItemDocOrigen_NovShr.ObjValorPro = ashrIdItemDocOrigen
            .ObjBaseDec.ObjValorPro = aobjNovedad.ObjBaseDec.ObjValorPro
            .ObjFactorDbl.ObjValorPro = aobjNovedad.ObjFactorDbl.ObjValorPro
            .ObjIdCuentaDb_NovStr.ObjValorPro = aobjNovedad.ObjIdCuentaCr_NovStr.ObjValorPro
            .ObjIdCuentaCr_NovStr.ObjValorPro = aobjNovedad.ObjIdCuentaDb_NovStr.ObjValorPro
            .ObjIdTipoNovedadByt.ObjValorPro = ClsOrionCop.FenuTipoNovContraria(
                    aobjNovedad.ObjIdTipoNovedadByt.ObjValorPro)
            .ObjValor_NovDec.ObjValorPro = aobjNovedad.ObjValor_NovDec.ObjValorPro
            .ObjIdAno_NovShr.ObjValorPro = aobjNovedad.ObjIdAno_NovShr.ObjValorPro
            .ObjIdServicio_NovShr.ObjValorPro = aobjNovedad.ObjIdServicio_NovShr.ObjValorPro
            ObjDebitos_ItemFactDec.ObjValorPro += .ObjValor_NovDec.ObjValorPro
            ObjFechaCancelacion_ItemFactDtm.ObjValorPro = GCDTMFECHANULA
        End With
        ColNovedades.Add(lobjNovedad)
    End Sub
    Friend Sub SReverseItemProgFact(adblIdCliente As Double)
        Dim lshrIdAno As Short = ObjIdAno_ServicioItemFactShr.ObjValorPro
        Dim lshrIdServicio As Short = ObjIdServicio_ItemFactShr.ObjValorPro
        Dim lstrIdPredio = ObjIdPredio_ItemFactStr.ToString()
        If String.IsNullOrEmpty(ObjIdPredio_ItemFactStr.ToString()) Then
            ClsOrionCop.SReverseItemProgFact(adblIdCliente, "", lshrIdAno, lshrIdServicio)
        Else
            ClsOrionCop.SReverseItemProgFact(0, lstrIdPredio, lshrIdAno, lshrIdServicio)
        End If
    End Sub
#End Region

#Region "Valores del servicio"
    ''' <summary>
    ''' Devuelve el valor del servicio sin tener en cuenta el IVA
    ''' </summary>
    ''' <returns></returns>
    Friend ReadOnly Property DecValorServicio() As Decimal
        Get
            Dim ldecValor As Decimal = ObjValor_ItemFactDec.ObjValorPro
            Dim ldblTasaIva As Double = ObjTarifaIva_ItemFactDbl.ObjValorPro
            Dim ldecValorServicio As Decimal = If(ldblTasaIva = 0, ldecValor,
                    Math.Round(ldecValor / (1 + ldblTasaIva)))
            Return ldecValorServicio
        End Get
    End Property

    Friend ReadOnly Property DecValorIntereses As Decimal
        Get
            Dim ldecValorInt = 0D
            Dim ldblTasaIva As Double = ObjTarifaIva_ItemFactDbl.ObjValorPro
            For Each lobjNov As ClsNovedad In ColNovedades
                If lobjNov.ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuDbInt Then
                    If ldblTasaIva > 0 Then
                        ldecValorInt += Math.Round(lobjNov.ObjValor_NovDec.ObjValorPro /
                                (1 + ldblTasaIva))
                    Else
                        ldecValorInt += lobjNov.ObjValor_NovDec.ObjValorPro
                    End If
                End If
            Next
            Return ldecValorInt
        End Get
    End Property

    ''' <summary>
    ''' Devuelve la base del IVA. Si la tasa de iva es cero, la base será cero
    ''' </summary>
    ''' <returns></returns>
    Friend ReadOnly Property DecBaseIvaServicio() As Decimal
        Get
            Dim ldecValor As Decimal = ObjValor_ItemFactDec.ObjValorPro
            Dim ldblTasaIva As Double = ObjTarifaIva_ItemFactDbl.ObjValorPro
            Dim ldecBaseIva As Decimal = If(ldblTasaIva = 0, 0,
                    Math.Round(ldecValor / (1 + ldblTasaIva)))
            Return ldecBaseIva
        End Get
    End Property

    ''' <summary>
    ''' Devuelve la valor del iva incluido en el valor del item
    ''' </summary>
    ''' <returns></returns>
    Friend ReadOnly Property DecIvaServicio() As Decimal
        Get
            Dim ldecValorAntes As Decimal, ldecIva = 0D
            Dim ldecValor As Decimal = ObjValor_ItemFactDec.ObjValorPro
            Dim ldblTasaIva As Double = ObjTarifaIva_ItemFactDbl.ObjValorPro
            If ldblTasaIva > 0 Then
                ldecValorAntes = Math.Round(ldecValor / (1 + ldblTasaIva))
                ldecIva = ldecValor - ldecValorAntes
            End If
            Return ldecIva
        End Get
    End Property

    ''' <summary>
    ''' Devuelve la base del iva  para aquellos intereses que causaron Iva. 
    ''' Solo se utiliza para la anulación de la factura.
    ''' </summary>
    ''' <returns></returns>
    Friend Function FdecBaseIvaIntSer() As Decimal
        Dim ldecBase = 0D
        If ObjTarifaIva_ItemFactDbl.ObjValorPro > 0 Then
            For Each lobjNov As ClsNovedad In ColNovedades
                If lobjNov.ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuDbInt Then
                    ldecBase += lobjNov.ObjValor_NovDec.ObjValorPro
                End If
            Next
        End If
        Return ldecBase
    End Function

    ''' <summary>
    ''' Devuelve el iva para aquellos intereses que causaron Iva. 
    ''' Solo se utiliza para la anulación de la factura.
    ''' </summary>
    ''' <returns></returns>
    Friend Function FdecIvaIntSer() As Decimal
        Dim ldecIvaInt = 0D
        If ObjTarifaIva_ItemFactDbl.ObjValorPro > 0 Then
            For Each lobjNov As ClsNovedad In ColNovedades
                If lobjNov.ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuDbIvaInt Then
                    ldecIvaInt += lobjNov.ObjValor_NovDec.ObjValorPro
                End If
            Next
        End If
        Return ldecIvaInt
    End Function
#End Region

#Region "Información deuda"
    ''' <summary>
    ''' Devuelve la deuda total del item (Debitos - Creditos)
    ''' </summary>
    Friend ReadOnly Property DecDeuda As Decimal
        Get
            Dim ldecDeudaItem As Decimal = ObjDebitos_ItemFactDec.ObjValorPro -
                    ObjCreditos_ItemFactDec.ObjValorPro
            Return ldecDeudaItem
        End Get
    End Property

    Friend ReadOnly Property DecDeudaSer As Decimal
        Get
            SCalculeDeudas()
            Return MdecDeudaSer
        End Get
    End Property

    Friend ReadOnly Property DecDeudaIvaSer As Decimal
        Get
            SCalculeDeudas()
            Return MdecDeudaIvaSer
        End Get
    End Property

    Friend ReadOnly Property DecDeudaInt As Decimal
        Get
            SCalculeDeudas()
            Return MdecDeudaInt
        End Get
    End Property

    Friend ReadOnly Property DecDeudaIvaInt As Decimal
        Get
            SCalculeDeudas()
            Return MdecDeudaIvaInt
        End Get
    End Property

    ' Deuda de capital
    ''' <summary>
    ''' Devuelve la deuda del servicio mas la deuda del IVA al servicio
    ''' </summary>
    ''' <returns></returns>
    Friend Function FdecDeudaCapital() As Decimal
        Dim ldecDeuda = DecDeudaSer() + DecDeudaIvaSer()
        Return ldecDeuda - MdecCrAplicadoCap
    End Function

    ' Deuda Int Mora
    ''' <summary>
    ''' Devuelve el total de los intereses debidos incluido el Iva a ellos,
    ''' el cual no se puede tocar a no ser que se anule la factura
    ''' </summary>
    ''' <returns></returns>
    Friend Function FdecDeudaIntTotal() As Decimal
        Dim ldecDeudaInt = DecDeudaInt() + DecDeudaIvaInt()
        Return ldecDeudaInt - MdecCrAplicadoInt
    End Function

    ' Total deuda IVA
    ''' <summary>
    ''' Devuelve la deuda del iva del servicio mas la deuda del iva de los intereses 
    ''' menos el iva al gasto
    ''' </summary>
    ''' <returns></returns>
    Friend Function FdecDeudaIvaTotal() As Decimal
        Dim ldecDeudaIva As Decimal = DecDeudaIvaSer() + DecDeudaIvaInt()
        Return ldecDeudaIva
    End Function

    ''' <summary>
    ''' A partir de las novedades del ítem de factura, calcula la deudas por servicio, 
    ''' por iva al servicio, por intereses y por iva a los intereses   
    ''' </summary>
    Private Sub SCalculeDeudas()
        If Not MblnCalculoDeudas Then
            Dim ldecDbsSer = 0D, ldecPagoCap = 0D, ldecDsctoCap = 0D
            Dim ldecDbsInt = 0D, ldecPagoInt = 0D, ldecDsctoInt = 0D
            Dim ldecDbsIvaInt = 0D, lenuTipoNov As EnuTipoNov
            Dim ldecDbsIvaSer = FdecDbsIvaSer(ldecDbsIvaInt)
            For Each lobjNov As ClsNovedad In ColNovedades
                lenuTipoNov = lobjNov.ObjIdTipoNovedadByt.ObjValorPro
                ' Capital
                Select Case lenuTipoNov
                    Case EnuTipoNov.EnuDbCap
                        ldecDbsSer += lobjNov.ObjValor_NovDec.ObjValorPro
                    Case EnuTipoNov.EnuCrPagoCap, EnuTipoNov.EnuCrAnApCap,
                            EnuTipoNov.EnuCrRetIva
                        ldecPagoCap += lobjNov.ObjValor_NovDec.ObjValorPro
                    Case EnuTipoNov.EnuCrDctoCap, EnuTipoNov.EnuCrRetFte,
                            EnuTipoNov.EnuCrRetIca, EnuTipoNov.EnuCrRetCre
                        ldecDsctoCap += lobjNov.ObjValor_NovDec.ObjValorPro
                    Case EnuTipoNov.EnuRDbCap
                        ldecDbsSer -= lobjNov.ObjValor_NovDec.ObjValorPro
                    Case EnuTipoNov.EnuRCrPagoCap, EnuTipoNov.EnuRCrAnApCap,
                            EnuTipoNov.EnuRCrRetIva
                        ldecPagoCap -= lobjNov.ObjValor_NovDec.ObjValorPro
                    Case EnuTipoNov.EnuRCrDctoCap, EnuTipoNov.EnuRCrRetFte,
                            EnuTipoNov.EnuRCrRetIca, EnuTipoNov.EnuRCrRetCre
                        ldecDsctoCap -= lobjNov.ObjValor_NovDec.ObjValorPro
                ' Intereses
                    Case EnuTipoNov.EnuDbInt
                        ldecDbsInt += lobjNov.ObjValor_NovDec.ObjValorPro
                    Case EnuTipoNov.EnuCrPagoInt, EnuTipoNov.EnuCrAnApInt
                        ldecPagoInt += lobjNov.ObjValor_NovDec.ObjValorPro
                    Case EnuTipoNov.EnuCrDctoInt
                        ldecDsctoInt += lobjNov.ObjValor_NovDec.ObjValorPro
                    Case EnuTipoNov.EnuRDbInt
                        ldecDbsInt -= lobjNov.ObjValor_NovDec.ObjValorPro
                    Case EnuTipoNov.EnuRCrPagoInt, EnuTipoNov.EnuRCrAnApInt
                        ldecPagoInt -= lobjNov.ObjValor_NovDec.ObjValorPro
                    Case EnuTipoNov.EnuRCrDctoInt
                        ldecDsctoInt -= lobjNov.ObjValor_NovDec.ObjValorPro
                End Select
            Next
            ' Deuda Intereses, Iva Intereses
            Dim ldecPagoIvaInt As Decimal
            If ldecPagoInt >= ldecDbsIvaInt Then
                ldecPagoIvaInt = ldecDbsIvaInt
                ldecPagoInt -= ldecPagoIvaInt
            Else
                MdecDeudaIvaInt = ldecDbsIvaInt - ldecPagoInt
                ldecPagoInt = 0
            End If
            MdecDeudaInt = ldecDbsInt - ldecPagoInt - ldecDsctoInt
            ' Deuda Servicio, Iva Servicio
            Dim ldecPagoIvaSer As Decimal
            If ldecPagoCap >= ldecDbsIvaSer Then
                ldecPagoIvaSer = ldecDbsIvaSer
                ldecPagoCap -= ldecPagoIvaSer
            Else
                MdecDeudaIvaSer = ldecDbsIvaSer - ldecPagoCap
                ldecPagoCap = 0
            End If
            MdecDeudaSer = ldecDbsSer - ldecPagoCap - ldecDsctoCap
            If DecDeuda <> MdecDeudaSer + MdecDeudaIvaSer + MdecDeudaInt + MdecDeudaIvaInt Then
                Throw New PanLException("Error en el cálculo de la deuda del ítem de factura.")
            End If
            MblnCalculoDeudas = True
        End If
    End Sub

    ''' <summary>
    ''' Devuelve los dbs de iva al servicio y por referencia devuelve los dbs de iva a los 
    ''' intereses, teniendo en cuenta los créditos de iva al gasto
    ''' </summary>
    ''' <param name="adecDbsIvaInt"></param>
    ''' <returns></returns>
    Private Function FdecDbsIvaSer(ByRef adecDbsIvaInt As Decimal) As Decimal
        Dim ldecDbsIvaSer = 0D, ldecDbsIvaInt = 0D, ldecIvalAlGasto = 0D,
                lenuTipoNov As EnuTipoNov
        For Each lobjNov As ClsNovedad In ColNovedades
            lenuTipoNov = lobjNov.ObjIdTipoNovedadByt.ObjValorPro
            Select Case lenuTipoNov
                Case EnuTipoNov.EnuDbIva
                    ldecDbsIvaSer += lobjNov.ObjValor_NovDec.ObjValorPro
                Case EnuTipoNov.EnuRDbIva
                    ldecDbsIvaSer -= lobjNov.ObjValor_NovDec.ObjValorPro
                Case EnuTipoNov.EnuDbIvaInt
                    ldecDbsIvaInt += lobjNov.ObjValor_NovDec.ObjValorPro
                Case EnuTipoNov.EnuRDbIvaInt
                    ldecDbsIvaInt -= lobjNov.ObjValor_NovDec.ObjValorPro
                Case EnuTipoNov.EnuCrIvaGas
                    ldecIvalAlGasto += lobjNov.ObjValor_NovDec.ObjValorPro
                Case EnuTipoNov.EnuRCrIvaGas
                    ldecIvalAlGasto -= lobjNov.ObjValor_NovDec.ObjValorPro
            End Select
        Next
        If ldecIvalAlGasto >= ldecDbsIvaSer Then
            ldecDbsIvaInt -= (ldecIvalAlGasto - ldecDbsIvaSer)
            ldecDbsIvaSer = 0
        Else
            ldecDbsIvaSer -= ldecIvalAlGasto
        End If
        adecDbsIvaInt = ldecDbsIvaInt
        Return ldecDbsIvaSer
    End Function
#End Region

#Region "Calculo intereses mora"
    ''' <summary>
    ''' Devuelve los intereses de mora por causar
    ''' </summary>
    ''' <param name="adtmFecha">Indica la fecha de causación</param>
    ''' <returns></returns>
    Friend Function FdecIntePorCausar(adtmFecha As Date) As Decimal
        Dim lentDiasMora = 0, ldecIntMora = 0D
        If FblnCausarInt(adtmFecha) Then
            ldecIntMora = FdecIntereseMora(adtmFecha, lentDiasMora)
        End If
        Return ldecIntMora
    End Function

    ''' <summary>
    ''' Causa los intereses de mora a la fecha del parametro "adtmFecha"
    ''' </summary>
    ''' <param name="adtmFecha">Fecha en que se estan causando los intereses</param>
    ''' <param name="adecValorCausado">Devuelve el valor de os intereses causados</param>
    ' Causa mora a item de factura en procesos FM
    Friend Sub SCauseMora(adtmFechaCausacion As Date, ByRef adecValorCausado As Decimal,
                ByRef aentDiasMora As Integer)
        Dim lblnCausarInt = FblnCausarInt(adtmFechaCausacion)
        If lblnCausarInt Then
            adecValorCausado = FdecIntereseMora(adtmFechaCausacion, aentDiasMora)
            If adecValorCausado > 0 Then
                SModifique()
                ObjDebitos_ItemFactDec.ObjValorPro += adecValorCausado
                ObjFechaCausoIntMora_Dtm.ObjValorPro = adtmFechaCausacion
            Else
                Dim lobjFac As ClsFactura = ObjPadre
                If (GobjParametros.FdblTasaMoraFecha(adtmFechaCausacion) = 0 OrElse
                        Not ObjServicio.FblnCausaMora) AndAlso
                        lobjFac.ObjFechaGraciaDtm.ObjValorPro < adtmFechaCausacion Then
                    SModifique()
                    ObjFechaCausoIntMora_Dtm.ObjValorPro = adtmFechaCausacion
                End If
            End If
        End If
    End Sub
    ' Llamado en procesos FM
    Private Function FdecIntereseMora(adtmFechaCalculo As Date, ByRef aentDiasMora As Integer) As Decimal
        Dim ldecVlrMora = 0D
        aentDiasMora = FentCantidadDiasMora(adtmFechaCalculo)
        If aentDiasMora > 0 Then
            Dim lentDiasAno As Integer = 365
            If adtmFechaCalculo.Year Mod 4 = 0 Then
                lentDiasAno += 1
            End If
            Dim ldblTasaMora As Double = GobjParametros.FdblTasaMoraFecha(adtmFechaCalculo)
            ldecVlrMora = FdecDeudaCapital() * (ldblTasaMora / lentDiasAno) * aentDiasMora
            If ObjServicio.ObjTarifaIvaDbl.ObjValorPro > 0 Then
                ldecVlrMora += ldecVlrMora * ObjServicio.ObjTarifaIvaDbl.ObjValorPro
            End If
            ldecVlrMora = ClsOrionCop.FdecValorMoraRedondeado(ldecVlrMora)
        End If
        Return ldecVlrMora
    End Function

    Private Function FblnCausarInt(adtmFechaCausacion As Date) As Boolean
        Dim lblnCausar As Boolean = ObjServicio.FblnCausaMora
        If lblnCausar Then
            Dim ldtmFechaCausa = adtmFechaCausacion
            lblnCausar = ldtmFechaCausa > ObjFechaGraciaIFDtm.ObjValorPro AndAlso
                    ldtmFechaCausa > ObjFechaVencimientoIFDtm.ObjValorPro AndAlso
                    ldtmFechaCausa > ObjFechaCausoIntMora_Dtm.ObjValorPro
            If lblnCausar Then
                Dim lenuModoCausa As EnuModoCausaMora =
                        ObjServicio.ObjModoCausaInteresesByt.ObjValorPro
                If GblnCausandoFM Then
                    lblnCausar = GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro
                    If Not lblnCausar Then
                        lblnCausar = lenuModoCausa = EnuModoCausaMora.EnuFinMes OrElse lenuModoCausa =
                            EnuModoCausaMora.EnuUltimoDia OrElse lenuModoCausa =
                            EnuModoCausaMora.EnuEnFecha
                    End If
                Else
                    lblnCausar = lenuModoCausa = EnuModoCausaMora.EnuEnFecha OrElse
                            lenuModoCausa = EnuModoCausaMora.EnuAlReciboCaja
                End If
            End If
        End If
        Return lblnCausar
    End Function
    ' Llamado en los procesos FM
    Private Function FentCantidadDiasMora(adtmFechaCalculo As Date) As Integer
        Dim ldtmFechaSinMora As Date = ObjFechaGraciaIFDtm.ObjValorPro
        Dim lentCantDias = 0
        If adtmFechaCalculo > ldtmFechaSinMora Then
            Dim ldtmFechaVence As Date = ObjFechaVencimientoIFDtm.ObjValorPro
            Dim ldtmFechaUltiCausMora As Date = ObjFechaCausoIntMora_Dtm.ObjValorPro
            Dim ldtmfechaIni As Date = If(ldtmFechaUltiCausMora = GCDTMFECHANULA OrElse
                    ldtmFechaVence >= ldtmFechaUltiCausMora, ldtmFechaVence, ldtmFechaUltiCausMora)
            lentCantDias = ClsPanorama.FentDiasEntreFechas(ldtmfechaIni, adtmFechaCalculo)
        End If
        Return lentCantDias
    End Function

    Friend Function FstrIdCtaMora() As String
        Dim lstrIdCtaMora = ObjServicio.ObjCodigoCuentaMoraStr.ObjValorPro
        Return lstrIdCtaMora
    End Function
#End Region

#Region "Retenciones y Descuento por Pronto Pago"
    ''' <summary>
    ''' Devuelve el valor que se debe aplicar al Itemd se Factura según los parametros establecidos
    ''' en la Copropiedad y en el Servicio
    ''' </summary>
    ''' <param name="aenuTipoRetencion"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function FdecValorRetencion(aenuTipoRetencion As EnuTipoDescuento,
            ByRef adecBaseRet As Decimal, ByRef adblTasaRet As Double)
        Dim ldecVlrRetencion = 0D
        adecBaseRet = 0
        adblTasaRet = 0
        Select Case aenuTipoRetencion
            Case EnuTipoDescuento.EnuReteFuente
                ldecVlrRetencion = FdecReteFuente(adecBaseRet, adblTasaRet)
            Case EnuTipoDescuento.EnuReteIca
                ldecVlrRetencion = FdecReteIca(adecBaseRet, adblTasaRet)
            Case EnuTipoDescuento.EnuReteCree
                ldecVlrRetencion = 0
            Case EnuTipoDescuento.EnuReteIva
                ldecVlrRetencion = FdecReteIva(adecBaseRet, adblTasaRet)
        End Select
        Return ldecVlrRetencion
    End Function
    ''' <summary>
    ''' Devuelve el valor del descuento por pronto pago a que tiene derecho en el supuesto caso
    ''' que se cancele antes de la fecha indicada. Se calcula al momento de generar la factura.
    ''' </summary>
    ''' <remarks>Se calcula al momento de generar la factura.</remarks>
    Friend Function FdecDsctoPPPosible() As Decimal
        Dim ldecDsctoPP = 0D
        If GobjParametros.ObjAnoActual.ObjTipoIncentivoByt.ObjValorPro =
                EnuTipoIncentivo.EnuDescuentoPP AndAlso Not IsNothing(ObjSector) Then
            Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.EnuUnico)
            Dim lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil,
                        ObjIdPredio_ItemFactStr.ObjValorPro}
            lobjPredio.SAbra(lobjValorLlave)
            Dim ldblIdCliente = MobjMiFactura.ObjIdCliente_FactDbl.ObjValorPro
            ldecDsctoPP = lobjPredio.FdecDesctoPP_Prop(ldblIdCliente,
                        ObjValor_ItemFactDec.ObjValorPro)
        End If
        Return ldecDsctoPP
    End Function
    Friend Function FdecDsctoPPAAplicar(adtmFechaAplicacion As Date)
        Dim ldecDsctoPPItem = 0D
        If ObjServicio.BlnEsCuotaAdministracion Then
            If adtmFechaAplicacion <= MobjMiFactura.ObjFechaDctoProntoPagoDtm.ObjValorPro Then
                ldecDsctoPPItem = FdecDsctoPPPosible()
            End If
        End If
        Return ldecDsctoPPItem
    End Function
    Private Function FdecReteFuente(ByRef adecBaseRet As Decimal, ByRef adblTasaRet As Double) As Decimal
        Dim ldecRetFte = 0D
        If MobjMiFactura.ObjClienteFactura.ObjEsAgenteReteFteBln.ObjValorPro Then
            adblTasaRet = ObjServicio.ObjTarifaRetFteDbl.ObjValorPro
            If adblTasaRet > 0 Then
                If Not FblnRetencionAplicada(EnuTipoDescuento.EnuReteFuente) Then
                    Dim ldecBaseMinRet As Decimal =
                            ObjServicio.ObjBaseMinimaReteFuenteDec.ObjValorPro
                    adecBaseRet = DecBaseIvaServicio()

                    If adecBaseRet >= ldecBaseMinRet AndAlso FdecDeudaCapital() >= adecBaseRet Then
                        ldecRetFte = adecBaseRet * adblTasaRet
                        ldecRetFte = Math.Round(ldecRetFte, 0)
                    End If
                End If
            End If
        End If
        Return ldecRetFte
    End Function
    Private Function FdecReteIca(ByRef adecBaseRet As Decimal, ByRef adblTasaRet As Double) As Decimal
        Dim ldecRetIca = 0D
        If MobjMiFactura.ObjClienteFactura.ObjRetieneIcaBln.ObjValorPro Then
            If Not FblnRetencionAplicada(EnuTipoDescuento.EnuReteIca) Then
                adblTasaRet = ObjServicio.ObjTarifaRetIcaDbl.ObjValorPro
                If adblTasaRet > 0 Then
                    Dim ldecBaseMinRet As Decimal = ObjServicio.ObjBaseMinimaReteIcaDec.ObjValorPro
                    adecBaseRet = DecValorServicio()

                    If adecBaseRet >= ldecBaseMinRet AndAlso FdecDeudaCapital() >= adecBaseRet Then
                        ldecRetIca = adecBaseRet * adblTasaRet
                        ldecRetIca = Math.Round(ldecRetIca, 0)
                    End If
                End If
            End If
        End If
        Return ldecRetIca
    End Function
    Private Function FdecReteIva(ByRef adecBaseRet As Decimal, ByRef adblTasaRet As Double) As Decimal
        Dim ldecReteIva = 0D
        If MobjMiFactura.ObjClienteFactura.ObjRetieneIvaBln.ObjValorPro Then
            If Not FblnRetencionAplicada(EnuTipoDescuento.EnuReteIva) Then
                adecBaseRet = DecIvaServicio()
                If adecBaseRet > 0 AndAlso FdecDeudaCapital() >= adecBaseRet Then
                    adblTasaRet = GobjParametros.ObjTarifaReteIvaDbl.ObjValorPro
                    If adblTasaRet > 0 Then
                        ldecReteIva = adblTasaRet * adecBaseRet
                        ldecReteIva = Math.Round(ldecReteIva, 0)
                    End If
                End If
            End If
        End If
        Return ldecReteIva
    End Function
#End Region

#Region "Datos EFactura"
    ''' <summary>
    ''' Devuelve la propiedad ItemReference del Json de la factura para Efactura
    ''' </summary>
    Friend ReadOnly Property StrItemRef As String
        Get
            Dim lstrItemRef As String = FstrComillas(ObjIdAno_ServicioItemFactShr.ToString &
                Format(ObjIdServicio_ItemFactShr.ObjValorPro, "00#"))
            Return lstrItemRef
        End Get
    End Property
    ''' <summary>
    ''' Devuelve la propiedad Name del Json de la factura para Efactura
    ''' </summary>
    Friend ReadOnly Property StrName As String
        Get
            Return FstrComillas(ObjDetalle_ItemFactStr.ToString)
        End Get
    End Property
    ''' <summary>
    ''' Devuelve la propiedad Price del Json de la factura para Efactura. Corresponde al precio del
    ''' servicio antes del IVA
    ''' </summary>
    Friend ReadOnly Property StrPrice As String
        Get
            Return Format(DecValorServicio(), "#0.00")
        End Get
    End Property
    ''' <summary>
    ''' Devuelve la propiedad LineExtensionAmount del Json de la factura para Efactura. 
    ''' Es igual a Cantidad * ValorAntesIva - Descuentos + cargos.
    ''' </summary>
    Friend ReadOnly Property StrLineExtensionAmount As String
        Get
            Dim ldecRes As Decimal = DecValorServicio()
            Return Format(ldecRes, "#0.00")
        End Get
    End Property
    ''' <summary>
    ''' Devuelve la propiedad LineTotalTaxes que en este caso es el valor del Iva
    ''' </summary>
    ''' <returns></returns>
    Friend ReadOnly Property StrLineTotalTaxes
        Get
            Return Format(DecIvaServicio, "#0.00")
        End Get
    End Property
    ''' <summary>
    ''' Devuelve la propiedad LineTotal que es la suma de StrLineExtensionAmount y StrLineTotalTaxes
    ''' </summary>
    ''' <returns></returns>
    Friend ReadOnly Property StrLineTotal As String
        Get
            Return Format(ObjValor_ItemFactDec.ObjValorPro, "#0.00")
        End Get
    End Property
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsCreditos_ItemFactDec
    Inherits clsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Creditos"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Creditos_ItemFactura"
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

Friend Class ClsDebitos_ItemFactDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Debitos"
    Private ReadOnly MobjPadre As ClsItemFactura = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Debitos_ItemFactura"
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

Friend Class ClsDetalle_ItemFactStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Detalle"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "DetalleItemsFact"
        HshrLongitud = 100
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, HshrLongitud, HblnEsRequerido)
        If HblnEsValido Then
            Dim lobjPadre As ClsItemFactura = ObjPadre
            If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = (HobjValorNew.ToString.Trim() = HobjValorOriginal.ToString.Trim())
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

Friend Class ClsEsExcluidoIva_ItemFactBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "EsExcluidoIva"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "EsExcluidoIva_ItemFact"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
        If HblnEsValido Then
            Dim lobjPadre As ClsItemFactura = ObjPadre
            If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            End If
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

Friend Class ClsFechaCancelacion_ItemFactDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaCancelacion"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "FechaCancelacion_ItemFact"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = False
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsItemFactura = ObjPadre
        HblnEsValido = True
        If TypeOf lobjPadre.ObjPadre Is ClsFactura Then
            Dim lobjFactura As ClsFactura = lobjPadre.ObjPadre
            Dim ldtmFechaMin As Date = GCDTMFECHANULA
            Dim ldtmFechaMax As Date = Date.Today
            If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                ldtmFechaMax = GCDTMFECHANULA
            ElseIf lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                Dim ldtmFechaIniPeriodoActual = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
                If lobjPadre.ObjDebitos_ItemFactDec.ObjValorPro = lobjPadre.ObjCreditos_ItemFactDec.ObjValorPro Then
                    If ldtmFechaIniPeriodoActual >= lobjFactura.ObjFechaFacturaDtm.ObjValorPro Then
                        ldtmFechaMin = ldtmFechaIniPeriodoActual
                    Else
                        ldtmFechaMin = lobjFactura.ObjFechaFacturaDtm.ObjValorPro
                    End If
                End If
                If Date.Today > GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo Then
                    ldtmFechaMax = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
                Else
                    ldtmFechaMax = Now
                End If
            End If
            HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax, BlnEsRequerido)
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

Friend Class ClsFechaCausoIntMora_Dtm
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsItemFactura = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaCausoMora"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Fecha Causo Mora"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim ldtmFechaMin As Date
        Dim ldtmFechaMax As Date
        Dim lobjMiFactura As ClsFactura = MobjPadre.ObjMiFactura
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando OrElse
                ClsOrionCop.BlnFacturando Then
            ldtmFechaMin = GCDTMFECHANULA
            ldtmFechaMax = GCDTMFECHANULA
            HblnEsRequerido = False
        ElseIf MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
            HblnEsRequerido = True
            Dim ldtmFechaIniPeriodoActual = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
            Dim ldtmFechaFinPeriodoActual = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
            If ldtmFechaIniPeriodoActual < lobjMiFactura.ObjFechaFacturaDtm.ObjValorPro Then
                If lobjMiFactura.ObjFechaFacturaDtm.ObjValorPro = ldtmFechaFinPeriodoActual Then
                    ldtmFechaMin = lobjMiFactura.ObjFechaFacturaDtm.ObjValorPro
                Else
                    ldtmFechaMin = lobjMiFactura.ObjFechaFacturaDtm.ObjValorPro
                    ldtmFechaMin = ldtmFechaMin.AddDays(1)
                End If
            Else
                ldtmFechaMin = ldtmFechaIniPeriodoActual
            End If
            If Date.Today > GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo Then
                ldtmFechaMax = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
                ldtmFechaMax = ldtmFechaMax.AddDays(1)
            Else
                ldtmFechaMax = Now
            End If
        Else
            ldtmFechaMax = HobjValorOriginal
            ldtmFechaMin = ldtmFechaMax
            HblnEsRequerido = Not (ldtmFechaMax = GCDTMFECHANULA)
        End If
        If ClsOrionCop.BlnProcesoEspecial Then
            ldtmFechaMin = GCDTMFECHANULA
            ldtmFechaMax = Now.AddDays(1)
            HblnEsRequerido = False
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

Friend Class ClsFechaGraciaIFDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaPagoSinMora"
    Private ReadOnly MobjPadre As ClsItemFactura = Nothing
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
            Dim lobjAbuelo As ClsFactura = MobjPadre.ObjPadre
            Dim ldtmFechaMin As Date = GCDTMFECHANULA
            Dim ldtmFechaMax As Date = Date.Today.AddYears(2)
            HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax,
                    BlnEsRequerido)
            HstrMens = String.Empty
            If HblnEsValido Then
                If HobjValorNew <> GCDTMFECHANULA Then
                    HblnEsValido = (HobjValorNew >= MobjPadre.ObjFechaVencimientoIFDtm.ObjValorPro)
                    If Not HblnEsValido Then
                        HstrMens = "La Fecha del Périodo de Gracia es anterior a la Fecha de vencimiento!"
                    End If
                End If
            Else
                If Not ClsOrionCop.BlnFacturando AndAlso
                        lobjAbuelo.ObjIdModoFacturacionByt.ObjValorPro =
                                EnuModoFacturacionDef.EnuManual Then
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

Friend Class ClsFechaVencimientoIFDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaVencimiento"
    Private ReadOnly MobjPadre As ClsItemFactura = Nothing
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
            Dim lobjAbuelo As ClsFactura = MobjPadre.ObjPadre
            Dim ldtmFechaMin As Date = lobjAbuelo.ObjFechaFacturaDtm.ObjValorPro
            Dim ldtmFechaMax As Date = ldtmFechaMin.AddYears(3)
            HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax,
                    BlnEsRequerido)
            If Not HblnEsValido Then
                If Not ClsOrionCop.BlnFacturando AndAlso
                        lobjAbuelo.ObjIdModoFacturacionByt.ObjValorPro =
                                EnuModoFacturacionDef.EnuManual Then
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
            MobjPadre.ObjFechaGraciaIFDtm.SValide()
        End If
    End Sub
End Class

Friend Class ClsIdAno_ServicioItemFactShr
    Inherits clsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdAno"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdAñoServicio_ItemFactura"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, Year(Date.MaxValue),
                BlnEsRequerido, EnuTipoValor)
        If HblnEsValido Then
            Dim lobjPadre As ClsItemFactura = ObjPadre
            If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                If HobjValorNew > 0 Then
                    HblnEsValido = GobjParametros.ColAnos.Contains(HobjValorNew.ToString)
                End If
            Else
                HblnEsValido = HobjValorNew = HobjValorOriginal
            End If
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If HblnEsValido Then
            Dim lobjPadre As ClsItemFactura = ObjPadre
            If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                lobjPadre.ObjIdServicio_ItemFactShr.SValide()
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

Friend Class ClsIdFactura_ItemFactEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdFactura"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdFactura_ItemFact"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 3
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsItemFactura = ObjPadre
        If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando OrElse
                ClsOrionCop.BlnFacturando Then
            Dim lobjFactura As ClsFactura = lobjPadre.ObjPadre
            HblnEsValido = (HobjValorNew = lobjFactura.ObjIdFacturaEnt.ObjValorPro)
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
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsIdItemFacturaShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdItemFactura"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdItemFactura"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 4
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue,
                BlnEsRequerido, EnuTipoValor)
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

Friend Class ClsIdPredio_ItemFactStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdPredio"
    Private ReadOnly MobjPadre As ClsItemFactura = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdPredioItemFactura"
        HenuTipoValor = EnuTipoValor.enuString
        HshrLongitud = 20
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = False
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                Dim lstrIdPredioAgru = String.Empty
                If HobjValorNew <> String.Empty Then
                    Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
                    Dim lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
                    lobjPredio.SAbra(lobjValorLlave)
                    HblnEsValido = lobjPredio.BlnExiste
                    If HblnEsValido Then
                        lstrIdPredioAgru = lobjPredio.ObjIdPredioAgrupadorStr.ObjValorPro
                    End If
                End If
                If Not ClsOrionCop.BlnProcesoEspecial Then
                    If HblnEsValido Then
                        If HobjValorNew <> "" Then
                            Dim lobjFactura As ClsFactura = MobjPadre.ObjPadre
                            HblnEsValido = lobjFactura.FblnPredioAgrupadorEsValido(lstrIdPredioAgru)
                            If Not HblnEsValido Then
                                HstrMens = "Todos los items de la Factura deben tener " &
                                    "relación con el mismo Predio Agrupador!"
                                SNotifiqueDatInv()
                            End If
                        End If
                    End If
                End If
            Else
                HblnEsValido = HobjValorNew = HobjValorOriginal
            End If
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

Friend Class ClsIdServicio_ItemFactShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdServicio"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdServicio_ItemFactura"
        HenuTipoValor = EnuTipoValor.enuUShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue,
                    BlnEsRequerido, EnuTipoValor)
        If HblnEsValido Then
            Dim lobjPadre As ClsItemFactura = ObjPadre
            If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                If lobjPadre.ObjIdAno_ServicioItemFactShr.BlnEsValido Then
                    Dim lshrIdano = lobjPadre.ObjIdAno_ServicioItemFactShr.ObjValorPro
                    Dim lcolServicios As Collection
                    If lshrIdano = 0 Then
                        lcolServicios = GobjParametros.ColServiciosPer
                    Else
                        If GobjParametros.ColAnos.Contains(lshrIdano.ToString) Then
                            Dim lobjAno As ClsAno = GobjParametros.ColAnos(lshrIdano.ToString)
                            lcolServicios = lobjAno.ColServiciosAno
                        Else
                            Throw New ErrorInesperadoPanLException("Año no existe")
                        End If
                    End If
                    Dim lstrKey = lobjPadre.ObjIdAno_ServicioItemFactShr.ToString & "," & HobjValorNew.ToString
                    HblnEsValido = lcolServicios.Contains(lstrKey)
                End If
            Else
                If Not ClsOrionCop.BlnProcesoEspecial Then
                    HblnEsValido = (HobjValorNew = HobjValorOriginal)
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
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsPeriodo_ItemFactStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Periodo"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Periodo_ItemFact"
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HshrLongitud = 6
        HblnEsRequerido = False
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoStringNumerico(HobjValorNew, ShrLongitud, ShrLongitud,
                    BlnEsRequerido)
        If HblnEsValido Then
            Dim lobjPadre As ClsItemFactura = ObjPadre
            If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                Dim lobjFactura As ClsFactura = lobjPadre.ObjPadre
                Dim ldtmFechaFact As Date = lobjFactura.ObjFechaFacturaDtm.ObjValorPro
                HblnEsValido = (HobjValorNew >= ClsOrionCop.FstrPeriodoDeFecha(ldtmFechaFact))
            Else
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
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

Friend Class ClsPrefijo_ItemFactStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Prefijo"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "PrefijoFactura"
        HshrLongitud = 5
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsLlave = True
        HbytPosicionLlave = 2
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then
            HobjValorNew = String.Empty
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            Dim lobjPadre As ClsItemFactura = ObjPadre
            If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando OrElse
                    ClsOrionCop.BlnFacturando Then
                Dim lobjFactura As ClsFactura = lobjPadre.ObjPadre
                HblnEsValido = (HobjValorNew = lobjFactura.ObjPrefijo_FactStr.ObjValorPro)
            Else
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
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

Friend Class ClsTarifaIva_ItemFactDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "TarifaIva"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "TarifaDelIva_ItemFactura"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, 0.5, BlnEsRequerido, HenuTipoValor)
        If HblnEsValido Then
            Dim lobjPadre As ClsItemFactura = ObjPadre
            If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
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
            Return Format(HobjValorPro, "p")
        End If
    End Function
End Class

''' <summary>
''' Guarda el valor del Item facturado, el cual tiene incluido el Iva
''' </summary>
''' <remarks></remarks>
Friend Class ClsValor_ItemFactDec
    Inherits clsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Valor"
    Private ReadOnly MobjPadre As clsItemFactura = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Valor Item Factura"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjValorIng As Object = HobjValorNew
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Decimal.MaxValue,
                BlnEsRequerido, HenuTipoValor)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = HobjValorNew - Int(HobjValorNew) = 0
                If Not HblnEsValido Then
                    HstrMens = "El Valor ingresado debe ser sin Centavos!"
                    SNotifiqueDatInv()
                End If
            Else
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            End If
        Else
            If HobjValorNew = 0 Then
                HblnEsValido = MobjPadre.ObjAnuladoBln.ObjValorPro
            End If
            If Not HblnEsValido Then
                If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                    HstrMens = "El valor ingresado, '" & lobjValorIng.ToString & ", no es valido"
                    SNotifiqueDatInv()
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
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return Format(HobjValorPro, "c")
        End If
    End Function
End Class
#End Region
Friend Class ClsReciboCaja
#Region "Definiciones"
    Inherits ClsCBObjetoPan
#Region "Variables y constantes"
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriRecibosCaja"
    ' Variables de modulo
    Private MobjClienteRec As ClsCliente = Nothing
    Private McolFrasVivas As Collection = Nothing
    Private MobjPredioAgrRec As ClsPredio = Nothing
    Private McolItemsRecCaja As Collection = Nothing
    Private MdtbItemsRecCaja As DataTable = Nothing
    Private McolMediosPago As Collection = Nothing
    Private MdtbMediosPago As DataTable = Nothing
    Private McolNovedades As Collection = Nothing
    Private MdtbNovedades As DataTable = Nothing
    Private MobjAnticipo As ClsAnticipo = Nothing
    Private MobjNotaReversionRC As ClsNotaReversionCr = Nothing
    Private MstrNroNotaRCr As String = String.Empty
    Private MdecInteresesMoraPorCausar As Decimal = 0
    Private MdecIntPorCausar As Decimal = 0
    Private MdecDeudaRC As Decimal = 0
    Private MarlPrediosAgrCliente As ArrayList = Nothing
    Private MdtbDescuentos As DataTable = Nothing
#End Region
#End Region

#Region "Constructores"
    ''' <summary>
    ''' Instancia un objeto Recibo de caja en modo único
    ''' </summary>
    Public Sub New()
        HobjPadre = Nothing
        HblnEsCreable = False
        HblnEsModificable = False
        HblnEsAnulable = False
        HblnEsSuprimible = False
        HenuTipoObjeto = EnuModoInstanciaObjDef.EnuUnico
        HcolTablas.Add(MCSTRNOMBRETABLA)
        HcolCamposSelect.Add({"*"})
    End Sub
    ''' <summary>
    ''' Instancia un objeto Recibo de Caja en modo navegable
    ''' </summary>
    Public Sub New(astrPref As String)
        HobjPadre = Nothing
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion
        lstrFiltro &= " AND " & ClsPrefijo_RecStr.SstrNombreCampoBd & " = '" & astrPref & "'"
        HcolFiltros.Add(lstrFiltro)
        Dim lstrCamposSelect = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsPrefijo_RecStr.SstrNombreCampoBd, ClsIdRecCajaEnt.SstrNombreCampoBd}
        HenuTipoObjeto = EnuModoInstanciaObjDef.EnuNavegable
        HblnEsModificable = False
        HblnEsSuprimible = False
        HcolTablas.Add(MCSTRNOMBRETABLA)
        HcolCamposSelect.Add(lstrCamposSelect)
    End Sub
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Administrador al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwObjeto">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Public Sub New(aobjPadre As ClsCliente, adrwObjeto As DataRow)
        HobjPadre = aobjPadre
        MobjClienteRec = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.EnuDeColeccion
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
    Friend Shared ReadOnly Property SstrNombreTabla As String
        Get
            Return MCSTRNOMBRETABLA
        End Get
    End Property
    Protected Overrides ReadOnly Property HenuIdClase As EnuIdClasesPanDef
        Get
            Return EnuIdClasesPanDef.EnuReciboCaja
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Recibo Caja"
        End Get
    End Property
    Friend Overrides ReadOnly Property HstrNombreObj As String
        Get
            Return Chr(34) & "Nro. " & StrNumeroRecCaja & Chr(34)
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjComentario_RecStr As New ClsComentario_RecStr(Me)
    Friend ReadOnly Property ObjCUDocStr As New ClsCUDocStr(Me)
    Friend ReadOnly Property ObjRCEnviadoMailBln As New ClsRCEnviadoMailBln(Me)
    Friend ReadOnly Property ObjFechaAnulacion_RecDtm As New ClsFechaAnulacion_RecDtm(Me)
    Friend ReadOnly Property ObjFechaRecDtm As New ClsFechaRecDtm(Me)
    Friend ReadOnly Property ObjIdAnticipo_RecEnt As New ClsIdAnticipo_RecEnt(Me)
    Friend ReadOnly Property ObjIdCarpeta_RecShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_RecShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdCliente_RecDbl As New ClsIdCliente_RecDbl(Me)
    Friend ReadOnly Property ObjIdNotasCrStr As New ClsIdNotasCrStr(Me)
    Friend ReadOnly Property ObjIdPredioAgrupador_RecStr As New ClsIdPredioAgrupador_RecStr(Me)
    Friend ReadOnly Property ObjIdRecCajaEnt As New ClsIdRecCajaEnt(Me)
    Friend ReadOnly Property ObjIdUsuario_RecStr As New ClsIdUsuarioStr(Me)
    Friend ReadOnly Property ObjPrefijo_RecStr As New ClsPrefijo_RecStr(Me)
    Friend ReadOnly Property ObjSaldo_RecDec As New ClsSaldo_RecDec(Me)
    Friend ReadOnly Property ObjServicios_RecStr As New ClsServicios_RecStr(Me)
    Friend ReadOnly Property ObjValor_RecDec As New ClsValor_RecDec(Me)
    Friend ReadOnly Property ObjValorAnticipoDec As New ClsValorAnticipoDec(Me)
    Friend ReadOnly Property ObjValorDeudaAlPagoDec As New ClsValorDeudaAlPagoDec(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjAnuladoBln)
                HcolPropiedades.Add(ObjFechaCreacionDtm)
                HcolPropiedades.Add(ObjIdUsuarioAnuloStr)
                HcolPropiedades.Add(ObjOrigenInstanciaStr)
                HcolPropiedades.Add(ObjOrigenInstanciaAnuloStr)
                HcolPropiedades.Add(ObjComentario_RecStr)
                HcolPropiedades.Add(ObjCUDocStr)
                HcolPropiedades.Add(ObjRCEnviadoMailBln)
                HcolPropiedades.Add(ObjFechaAnulacion_RecDtm)
                HcolPropiedades.Add(ObjFechaRecDtm)
                HcolPropiedades.Add(ObjIdAnticipo_RecEnt)
                HcolPropiedades.Add(ObjIdCarpeta_RecShr)
                HcolPropiedades.Add(ObjIdCentroUtil_RecShr)
                HcolPropiedades.Add(ObjIdCliente_RecDbl)
                HcolPropiedades.Add(ObjIdNotasCrStr)
                HcolPropiedades.Add(ObjIdPredioAgrupador_RecStr)
                HcolPropiedades.Add(ObjIdRecCajaEnt)
                HcolPropiedades.Add(ObjIdUsuario_RecStr)
                HcolPropiedades.Add(ObjPrefijo_RecStr)
                HcolPropiedades.Add(ObjSaldo_RecDec)
                HcolPropiedades.Add(ObjServicios_RecStr)
                HcolPropiedades.Add(ObjValor_RecDec)
                HcolPropiedades.Add(ObjValorAnticipoDec)
                HcolPropiedades.Add(ObjValorDeudaAlPagoDec)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    Friend Property BlnEsSoloAnticipo As Boolean = False
    Friend ReadOnly Property StrNumeroRecCaja As String
        Get
            Dim lstrNroRecCaja = String.Empty
            If EnuEstadoActualizacion <> EnuEstadoObjetoDef.EnuCreando Then
                lstrNroRecCaja = ObjPrefijo_RecStr.ObjValorPro
                If lstrNroRecCaja <> "" Then
                    lstrNroRecCaja &= "-"
                End If
                lstrNroRecCaja &= ObjIdRecCajaEnt.ToString
            End If
            Return lstrNroRecCaja
        End Get
    End Property
    Friend Property ObjClienteRecibo As ClsCliente
        Get
            If IsNothing(MobjClienteRec) Then
                Dim lobjValorLlave As Object() = {ObjIdCarpeta_RecShr.ObjValorPro,
                        ObjIdCentroUtil_RecShr.ObjValorPro, ObjIdCliente_RecDbl.ObjValorPro}
                MobjClienteRec = New ClsCliente(EnuModoInstanciaObjDef.EnuUnico)
                MobjClienteRec.SAbra(lobjValorLlave)
            End If
            Return MobjClienteRec
        End Get
        Set(value As ClsCliente)
            MobjClienteRec = value
            HobjPadre = MobjClienteRec
        End Set
    End Property
    Friend ReadOnly Property ObjPredioAgrRecCaja As ClsPredio
        Get
            If IsNothing(MobjPredioAgrRec) Then
                If Not IsNothing(ObjIdPredioAgrupador_RecStr.ObjValorPro) AndAlso
                        ObjIdPredioAgrupador_RecStr.ToString.Length > 0 Then
                    Dim lstrIdPreAgr = ObjIdPredioAgrupador_RecStr.ToString.Split(",")(0)
                    MobjPredioAgrRec = New ClsPredio(EnuModoInstanciaObjDef.EnuUnico)
                    Dim lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrIdPreAgr}
                    MobjPredioAgrRec.SAbra(lobjValorLlave)
                    If Not MobjPredioAgrRec.BlnExiste Then
                        MobjPredioAgrRec = Nothing
                    End If
                End If
            End If
            Return MobjPredioAgrRec
        End Get
    End Property
#End Region
#End Region

#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SVacie()
        MyBase.SVacie()
        BlnEsSoloAnticipo = False
        MdtbItemsRecCaja = Nothing
        McolItemsRecCaja = Nothing
        MobjClienteRec = Nothing
        MobjPredioAgrRec = Nothing
        MdtbMediosPago = Nothing
        McolMediosPago = Nothing
        MdtbNovedades = Nothing
        McolNovedades = Nothing
        MobjAnticipo = Nothing
        MobjNotaReversionRC = Nothing
        MstrNroNotaRCr = String.Empty
        McolFrasVivas = Nothing
        MdecIntPorCausar = 0
        MdecDeudaRC = 0
        MarlPrediosAgrCliente = Nothing
        MdtbDescuentos = Nothing
        MdecInteresesMoraPorCausar = 0
        MdecDeuda = 0
    End Sub
    Protected Overrides Sub SInicialiceObj()
        ObjAnuladoBln.ObjValorPro = False
        ObjFechaAnulacion_RecDtm.ObjValorPro = GCDTMFECHANULA
        ObjFechaCreacionDtm.ObjValorPro = Date.Now
        ObjIdCarpeta_RecShr.ObjValorPro = GshrIdCarpeta
        ObjIdCentroUtil_RecShr.ObjValorPro = GshrIdCentroUtil
        ObjOrigenInstanciaStr.ObjValorPro = GstrOrigenActual
        ObjIdUsuario_RecStr.ObjValorPro = GstrIdUsuario
        ObjIdUsuarioAnuloStr.ObjValorPro = String.Empty
        ObjOrigenInstanciaAnuloStr.ObjValorPro = String.Empty
        ObjComentario_RecStr.ObjValorPro = String.Empty
        Dim lstrPrefijo = GobjParametros.FstrPrefijoDoc(EnuTipoDocOri.EnuReciboCaja)
        If IsNothing(lstrPrefijo) Then lstrPrefijo = String.Empty
        ObjPrefijo_RecStr.ObjValorPro = lstrPrefijo
        ObjIdRecCajaEnt.ObjValorPro = 0
        ObjCUDocStr.ObjValorPro = String.Empty
        ObjIdNotasCrStr.ObjValorPro = String.Empty
    End Sub
    Public Overrides Function FblnEsAnulable() As Boolean
        Dim lblnEsAnulable = BlnEsAnulable AndAlso BlnExiste
        If lblnEsAnulable Then
            If GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro Then
                lblnEsAnulable = Date.Today <=
                        GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
            End If
        End If
        If lblnEsAnulable Then
            Dim lstrPeriodoRec = ClsPanorama.FstrPeriodo(ObjFechaRecDtm.ObjValorPro)
            Dim lstrPeriodoActual = GobjParametros.ObjAnoActual.StrIdPeriodoActual
            lblnEsAnulable = (lstrPeriodoRec = lstrPeriodoActual)
        End If
        If lblnEsAnulable Then
            If Not IsNothing(ObjAnticipo) Then
                lblnEsAnulable = (ObjAnticipo.DecAnticipoReintegrado = 0)
            End If
        End If
        If lblnEsAnulable Then
            lblnEsAnulable = FblnFechaDocEsPeriodoActual(ObjFechaRecDtm.ObjValorPro)
        End If
        Return lblnEsAnulable
    End Function
    Friend Overrides Function FblnEsCreable() As Boolean
        Dim lstrMens = String.Empty
        Dim lblnEsCreable = FblnNotificaOk(EnuIdMens.EnuNoCreable)
        If Not lblnEsCreable Then
            ClsOrionCop.FblnPuedeCrear(EnuTipoDocOri.EnuReciboCaja, False,
                    lstrMens)
            SLevanteEventoNot(lstrMens, "", EnuIdMens.EnuNoCreable,
                    EnuSeveridadNot.EnuFalta)
        End If
        Return lblnEsCreable
    End Function
    Protected Overrides Function SAnuleEnObj() As Boolean
        Dim lblnAnulado = FblnEsAnulable()
        If lblnAnulado Then
            ObjAnuladoBln.ObjValorPro = True
            ObjIdUsuarioAnuloStr.ObjValorPro = GstrIdUsuario
            ObjOrigenInstanciaAnuloStr.ObjValorPro = GstrOrigenActual
            If GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro Then
                ObjFechaAnulacion_RecDtm.ObjValorPro = Now
            Else
                Dim ldtmFecFinPer As Date =
                            GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
                If ldtmFecFinPer < Date.Today Then
                    ObjFechaAnulacion_RecDtm.ObjValorPro = ldtmFecFinPer
                Else
                    ObjFechaAnulacion_RecDtm.ObjValorPro = Now
                End If
            End If
            If lblnAnulado Then
                ' Nota Reversion crédito
                SGenereNotaRevCr()
                ' Reversar notas Cr
                If Not String.IsNullOrEmpty(ObjIdNotasCrStr.ToString) Then
                    SReverseNotasCr(ObjFechaAnulacion_RecDtm.ObjValorPro)
                End If
            End If
        End If
        Return lblnAnulado
    End Function
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        Dim lblnNoHayError = False
        GobjPanDat.SControleProcesoObj(True)
        Try
            GobjPanDat.SInicialiceTransaccion()
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
                ObjSaldo_RecDec.ObjValorPro = FdecSaldoInformar()
                SNumereObj()
                Dim lcolFrasVivas = FcolFrasVivas()
                SGenereNotasCR()
                SGenereItemsRecCaja()
                SAplicaItemsRecCaja()
                ClsPanorama.SActualiceCol(McolItemsRecCaja)
                ClsPanorama.SActualiceCol(McolMediosPago)
                If Not IsNothing(lcolFrasVivas) Then
                    ClsPanorama.SActualiceCol(lcolFrasVivas)
                End If
                If Not IsNothing(MobjAnticipo) Then
                    MobjAnticipo.SActualice(ablnExigeRequeridos)
                    ObjIdAnticipo_RecEnt.ObjValorPro = MobjAnticipo.ObjIdAnticipoEnt.ObjValorPro
                Else
                    ObjIdAnticipo_RecEnt.ObjValorPro = 0
                End If
                If Not FblnEsIntegroRC() Then
                    Throw New ErrorInesperadoPanLException("Problema de Integridad. El Recibo no fue guardado!")
                End If
                ObjFechaCreacionDtm.ObjValorPro = Date.Now
                MyBase.SActualice(ablnExigeRequeridos)
            Else
                ClsPanorama.SActualiceCol(ColItemsRecCaja)
                ClsPanorama.SActualiceCol(ColMediosPago)
                MyBase.SActualice(ablnExigeRequeridos)
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
            If lblnNoHayError Then
                GobjPanDat.SConfirmeTransaccion()
                GobjPanDat.SControleProcesoObj(False)
                ObjClienteRecibo.SRefresqueObj()
            Else
                MdtbItemsRecCaja = Nothing
                McolItemsRecCaja = Nothing
                MdtbMediosPago = Nothing
                McolMediosPago = Nothing
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            End If
            McolFrasVivas = Nothing
        End Try
    End Sub
    Friend Overrides Sub SRefresqueObj()
        If Not IsNothing(MobjClienteRec) AndAlso MobjClienteRec.BlnExiste Then
            MobjClienteRec.SRefresqueObj()
        End If
        MyBase.SRefresqueObj()
    End Sub
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Dim lstrIdIbjeto = ObjIdRecCajaEnt.ToString
            If ObjPrefijo_RecStr.ToString.Length > 0 Then
                lstrIdIbjeto = ObjPrefijo_RecStr.ObjValorPro & "-" & lstrIdIbjeto
            End If
            Return lstrIdIbjeto
        End Get
    End Property
#End Region

#Region "Nueva estrategia: Cada que cambia una propiedad que influye en la deuda ésta se verifica."
    Private MdecDeuda As Decimal
    Friend Sub SVerifiqueDeuda()
        MdecInteresesMoraPorCausar = 0
        MdecDeuda = 0
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando AndAlso FblnDatoDeudorOk() AndAlso
                ObjServicios_RecStr.BlnEsValido Then
            Dim lstrIdPrediosAgrupadores As String() =
                    ObjIdPredioAgrupador_RecStr.ToString().Split(",")
            Dim lstrServicios As String() = ObjServicios_RecStr.ToString().Split(",")
            MdecDeuda = ObjClienteRecibo.FdecDeuda(lstrIdPrediosAgrupadores, lstrServicios)
            If MdecDeuda > 0 Then
                Dim ldtmfechaRecCaja As Date = ObjFechaRecDtm.ObjValorPro
                MdecInteresesMoraPorCausar = ObjClienteRecibo.FdecIntMoraPorCausar(
                        lstrIdPrediosAgrupadores, ObjFechaRecDtm.ObjValorPro)
            End If
            SCargueDsctos()
        End If
    End Sub

    Friend ReadOnly Property DecInteresesPorCausar As Decimal
        Get
            Return MdecInteresesMoraPorCausar
        End Get
    End Property

    Friend ReadOnly Property DecDeuda As Decimal
        Get
            Return MdecDeuda
        End Get
    End Property

    Private Function FblnDatoDeudorOk() As Boolean
        Dim lblnDatOk = ObjFechaRecDtm.BlnEsValido AndAlso ObjIdCliente_RecDbl.BlnEsValido AndAlso
            ObjIdPredioAgrupador_RecStr.BlnEsValido
        Return lblnDatOk
    End Function

    Friend Sub SEstablezcaFechaRec()
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
            With GobjParametros
                If .ObjExigeFechaHoyDocsBln.ObjValorPro Then
                    ObjFechaRecDtm.ObjValorPro = Date.Today
                Else
                    If .ObjAnoActual.StrIdPeriodoActual <
                            ClsOrionCop.FstrPeriodoDeFecha(Date.Today) Then
                        ObjFechaRecDtm.ObjValorPro =
                            .ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
                    Else
                        ObjFechaRecDtm.ObjValorPro = Date.Today
                    End If
                End If
            End With
        End If
    End Sub
#End Region

#Region "Deuda, saldos y anticipo al crear"
    Friend Function FdecDeudaRC() As Decimal
        If ObjIdCliente_RecDbl.BlnEsValido AndAlso ObjIdPredioAgrupador_RecStr.BlnEsValido Then
            Static lstrServicio As String = String.Empty
            Static lstrIdPredAgr As String = String.Empty
            Dim lblnCalcular = MdecDeudaRC = 0
            If Not lblnCalcular Then
                lblnCalcular = lstrIdPredAgr <> ObjIdPredioAgrupador_RecStr.ToString() OrElse
                    lstrServicio <> ObjServicios_RecStr.ToString()
            End If
            If lblnCalcular Then
                lstrIdPredAgr = ObjIdPredioAgrupador_RecStr.ToString()
                lstrServicio = ObjServicios_RecStr.ToString()
                Dim lstrIdPrediosAgr As String() = lstrIdPredAgr.Split(",")
                Dim lstrServicios As String() = lstrServicio.Split(",")
                MdecDeudaRC = ObjClienteRecibo.FdecDeuda(lstrIdPrediosAgr, lstrServicios)
                If MdecDeudaRC = 0.0 AndAlso Not BlnEsSoloAnticipo Then
                    SLevanteEventoNot("No hay deudas pendientes de Pago!", "", 0,
                            EnuSeveridadNot.EnuInformacion)
                End If
            End If
        Else
            MdecDeudaRC = 0.0
        End If
        Return MdecDeudaRC
    End Function
    Friend Function FdecSaldoRC() As Decimal
        Dim ldecSaldo As Decimal
        If BlnEsSoloAnticipo Then
            ldecSaldo = 0
        Else
            Dim ldecDsctos As Decimal = DecTotalDsctos
            Dim ldecPago As Decimal = ObjValor_RecDec.ObjValorPro
            ldecSaldo = MdecDeuda - ldecDsctos - ldecPago
            If ldecSaldo < 0 Then
                ldecSaldo = 0
            End If
        End If
        Return ldecSaldo
    End Function
    Private Function FdecSaldoInformar() As Decimal
        Dim ldecSaldo As Decimal
        If GobjParametros.ObjInformaSaldoTotalDespuesRCBln.ObjValorPro Then
            Dim lstrIdPrediosAgr = ObjIdPredioAgrupador_RecStr.ToString().Split(",")
            Dim lstrServicios = {"A"}
            Dim ldecDeudaTotal = ObjClienteRecibo.FdecDeuda(lstrIdPrediosAgr, lstrServicios)
            Dim ldecDsctos As Decimal = DecTotalDsctos
            Dim ldecPago As Decimal = ObjValor_RecDec.ObjValorPro
            ldecSaldo = If(BlnEsSoloAnticipo, ldecDeudaTotal, ldecDeudaTotal - ldecDsctos -
                        ldecPago)
            If ldecSaldo < 0 Then
                ldecSaldo = 0
            End If
        Else
            ldecSaldo = FdecSaldoRC()
        End If
        Return ldecSaldo
    End Function
    Friend Function FdecAnticipo() As Decimal
        Dim ldecAnticipo As Decimal
        Dim ldecDeudaRC As Decimal = FdecDeudaRC()
        Dim ldecDsctos As Decimal = DecTotalDsctos
        Dim ldecPago As Decimal = ObjValor_RecDec.ObjValorPro
        Dim ldecSaldo = ldecDeudaRC - ldecDsctos - ldecPago
        If ldecSaldo < 0 Then
            ldecAnticipo = -ldecSaldo
        Else
            ldecAnticipo = 0
        End If
        ObjValorAnticipoDec.ObjValorPro = ldecAnticipo
        Return ldecAnticipo
    End Function
#End Region

#Region "Procedimientos del objeto"
    Private Sub SNumereObj()
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
            Dim lentIdRec As Integer
            Dim lstrPrefijo As String = ObjPrefijo_RecStr.ObjValorPro
            If IsNothing(lstrPrefijo) Then lstrPrefijo = String.Empty
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsPrefijo_RecStr.SstrNombreCampoBd & " = '" & lstrPrefijo & "'"
            lentIdRec = ClsPanorama.FobjUltimaIdNumericaObjeto(SstrNombreTabla,
                ClsIdRecCajaEnt.SstrNombreCampoBd, ObjIdRecCajaEnt.EnuTipoValor,
                lstrFiltro)
            If lentIdRec < GobjParametros.FentNumeracionInicialDoc(EnuTipoDocOri.EnuReciboCaja) Then
                lentIdRec = GobjParametros.FentNumeracionInicialDoc(EnuTipoDocOri.EnuReciboCaja)
            End If
            lentIdRec += 1
            ObjPrefijo_RecStr.ObjValorPro = lstrPrefijo
            ObjIdRecCajaEnt.ObjValorPro = lentIdRec
            For Each lobjMedioPago As ClsMedioPago In McolMediosPago
                lobjMedioPago.ObjPrefijo_MedPagoStr.ObjValorPro = lstrPrefijo
                lobjMedioPago.ObjIdRecCaja_MedPagoEnt.ObjValorPro = lentIdRec
            Next
        End If
    End Sub
    ''' <summary>
    ''' Causa intereses de mora al cliente y al predio agrupador del recibo.
    ''' </summary>
    ''' <returns>Devuelve el valor de los intereses de mora causados.</returns>
    ''' <remarks></remarks>
    Friend Function SCauseMoraClienteRC(ByRef astrMens As String) As Decimal
        Dim ldecIntMoraCausados = 0D
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
            Try
                If Not IsNothing(MobjClienteRec) AndAlso
                        ObjIdPredioAgrupador_RecStr.BlnEsValido AndAlso
                        ObjFechaRecDtm.BlnEsValido Then
                    Dim lstrPreAgupadores = ObjIdPredioAgrupador_RecStr.ToString.Split(",")
                    If Not IsNothing(lstrPreAgupadores) Then
                        ldecIntMoraCausados = MobjClienteRec.SCauseMora(lstrPreAgupadores,
                            ObjFechaRecDtm.ObjValorPro, astrMens)
                        SVerifiqueDeuda()
                    End If
                End If
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
            End Try
        End If
        MdecInteresesMoraPorCausar = 0
        Return ldecIntMoraCausados
    End Function
    ''' <summary>
    ''' Devuelve un string con la id de los predios agrupadores a los cuales se les aplico el RC
    ''' separados por comas
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function FstrIdPrediosAgrDelRC() As String
        Dim lstrIdPrediosAgr As String = String.Empty
        Dim lstrPredioAgr = String.Empty, lstrPrefFac = String.Empty, lentIdFac = 0
        Dim lobjFact As New ClsFactura()
        For Each lobjItemRC As ClsItemRecCaja In ColItemsRecCaja
            If lobjItemRC.ObjIdTipoItemRecByt.ObjValorPro <> EnuTipoItemRecCajaDef.EnuAnticipo Then
                With lobjItemRC
                    lstrPrefFac = .ObjPrefijoFact_ItemRecStr.ObjValorPro
                    lentIdFac = .ObjIdFactura_ItemRecEnt.ObjValorPro
                End With
                lobjFact.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefFac, lentIdFac})
                lstrPredioAgr = Trim(lobjFact.ObjIdPredioAgrupador_FacStr.ObjValorPro)
                If Not lstrIdPrediosAgr.Contains(lstrPredioAgr) Then
                    lstrIdPrediosAgr += lstrPredioAgr & ","
                End If
            Else
                If Not IsNothing(ObjAnticipo) Then
                    lstrPredioAgr = Trim(ObjAnticipo.ObjIdPredioAgrupador_AntStr.ObjValorPro)
                    If Not lstrIdPrediosAgr.Contains(lstrPredioAgr) Then
                        lstrIdPrediosAgr += lstrPredioAgr & ","
                    End If
                End If
            End If
        Next
        If Not IsNothing(ObjAnticipo) AndAlso lstrIdPrediosAgr.EndsWith(",") Then
            lstrIdPrediosAgr = lstrIdPrediosAgr.Substring(0, lstrIdPrediosAgr.Length - 1)
        End If
        Return lstrIdPrediosAgr
    End Function
    Friend Function FarlCorreosRec() As ArrayList
        Dim larlListaCorreos As New ArrayList
        Dim lstrCorreoCli As String, lstrCorreoPredio As String
        If ObjClienteRecibo.ObjRecibeDocsPorEmailBln.ObjValorPro Then
            lstrCorreoCli = ObjClienteRecibo.ObjEmailStr.ToString
            If Not String.IsNullOrEmpty(lstrCorreoCli) Then
                larlListaCorreos.Add(lstrCorreoCli)
            End If
        End If
        If Not IsNothing(ObjPredioAgrRecCaja) Then
            lstrCorreoPredio = ObjPredioAgrRecCaja.ObjEmailAdiStr.ToString
            If Not String.IsNullOrEmpty(lstrCorreoPredio) Then
                larlListaCorreos.Add(lstrCorreoPredio)
            End If
        End If
        Return larlListaCorreos
    End Function
    Friend Function FdtbRecsCajaUltimoMes() As DataTable
        Dim lstrTablaPri = SstrNombreTabla
        Dim lstrTablaSec = ClsCliente.SstrNombreTabla
        Dim lstrCampSelPri As String() = {ClsFechaRecDtm.SstrNombreCampoBd,
                ClsPrefijo_RecStr.SstrNombreCampoBd, ClsIdRecCajaEnt.SstrNombreCampoBd,
                ClsIdPredioAgrupador_RecStr.SstrNombreCampoBd, ClsIdCliente_RecDbl.SstrNombreCampoBd,
                ClsValor_RecDec.SstrNombreCampoBd, ClsAnuladoBln.SstrNombreCampoBd}
        Dim lstrCampSelSec As String() = {ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCampRelPri As String() = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsIdCliente_RecDbl.SstrNombreCampoBd}
        Dim lstrCampRelSec As String() = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrFechaIni = "'" & ClsPanoramaDat.FstrFechaNormalizada(Date.Today.AddMonths(-2)) & "'"
        Dim lstrFechaFin = "'" & ClsPanoramaDat.FstrFechaNormalizada(Date.Today) & "'"
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion_Pri & " AND " &
                ClsFechaRecDtm.SstrNombreCampoBd & " BETWEEN " & lstrFechaIni & " AND " &
                lstrFechaFin
        Dim lstrOrden As String(,) = {{ClsFechaRecDtm.SstrNombreCampoBd, "DESC"},
                {ClsPrefijo_RecStr.SstrNombreCampoBd, "ASC"},
                {ClsIdRecCajaEnt.SstrNombreCampoBd, "DESC"}}
        Dim ldtbRecsCaja = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCampSelPri, lstrTablaSec,
                lstrCampSelSec, lstrCampRelPri, lstrCampRelSec, lstrOrden, True, lstrFiltro, {})
        Return ldtbRecsCaja
    End Function
#End Region

#Region "Manejo Items Recibo Caja"
#Region "Genera Items Rec Caja"
    Private Sub SGenereItemsRecCaja()
        SCargueDtbItemsRecCaja()
        SGenereItemsRecPago()
        SGenereItemsRecAnticipo()
    End Sub

    Private Sub SGenereItemsRecPago()
        Dim ldecPagoTotal As Decimal = ObjValor_RecDec.ObjValorPro -
                ObjValorAnticipoDec.ObjValorPro
        If ldecPagoTotal > 0 Then
            Dim ldecPorAplicar = ldecPagoTotal
            Dim ldecAAplicar As Decimal, ldecMoraFac As Decimal
            Dim lblnAplicaAFra As Boolean
            McolItemsRecCaja = ColItemsRecCaja
            ' Items Aplican a Mora
            Dim lstrIdPreAgrFac As String
            For Each lobjFactura As ClsFactura In McolFrasVivas
                lstrIdPreAgrFac = lobjFactura.ObjIdPredioAgrupador_FacStr.ObjValorPro
                lblnAplicaAFra = ObjIdPredioAgrupador_RecStr.ToString.Split(",").Contains(lstrIdPreAgrFac)
                If lblnAplicaAFra Then
                    ldecAAplicar = 0
                    ldecMoraFac = lobjFactura.FdecDeudaIntTotal
                    If ldecPorAplicar > 0 AndAlso ldecMoraFac > 0 Then
                        ldecAAplicar = If(ldecMoraFac <= ldecPorAplicar, ldecMoraFac, ldecPorAplicar)
                        SGenereItemsRecMedPagoMora(lobjFactura, ldecAAplicar)
                        ldecPorAplicar -= ldecAAplicar
                    End If
                End If
            Next
            ' Aplicar a Capital
            If ldecPorAplicar > 0 Then
                Dim lstrPredsAgru = ObjIdPredioAgrupador_RecStr.ToString().Split(",")
                Dim lstrServicios = ObjServicios_RecStr.ToString().Split(",")
                For Each lobjFactura As ClsFactura In McolFrasVivas
                    Dim ldecDeudaCapital = lobjFactura.FdecDeudaCapitalSer(lstrServicios)
                    lstrIdPreAgrFac = lobjFactura.ObjIdPredioAgrupador_FacStr.ObjValorPro
                    lblnAplicaAFra = lstrPredsAgru.Contains(lstrIdPreAgrFac)
                    If lblnAplicaAFra Then
                        lblnAplicaAFra = ldecDeudaCapital > 0
                    End If
                    If lblnAplicaAFra Then
                        ldecAAplicar = 0
                        If ldecPorAplicar > 0 AndAlso ldecDeudaCapital > 0 Then
                            ldecAAplicar = If(ldecDeudaCapital <= ldecPorAplicar, ldecDeudaCapital,
                                    ldecPorAplicar)
                            SGenereItemsRecPagoFactCap(lobjFactura, ldecAAplicar)
                            ldecPorAplicar -= ldecAAplicar
                        End If
                        If ldecPorAplicar = 0 Then Exit For
                    End If
                Next
            End If
            If ldecPorAplicar <> 0 Then
                Throw New ErrorInesperadoPanLException("Valor del pago no fue aplicado adecuadamente!")
            End If
        End If
    End Sub

    Private Sub SGenereItemsRecPagoFactCap(aobjFactura As ClsFactura, adecValor As Decimal)
        Dim lstrServicios As String() = ObjServicios_RecStr.ToString().Split(",")
        Dim ldecPorAplicar = adecValor, ldecAAplicar As Decimal, lblnAplicar As Boolean
        For Each lobjItemFac As ClsItemFactura In aobjFactura.ColItemsFactura
            lblnAplicar = False
            ldecAAplicar = 0
            If lobjItemFac.ObjIdAno_ServicioItemFactShr.ObjValorPro > 0 Then
                If lstrServicios.Contains("0") OrElse lstrServicios.Contains("A") Then
                    lblnAplicar = True
                End If
            ElseIf lstrServicios.Contains(lobjItemFac.ObjIdServicio_ItemFactShr.ToString()) OrElse
                    lstrServicios.Contains("A") Then
                lblnAplicar = True
            End If
            If lblnAplicar Then
                ldecAAplicar = If(lobjItemFac.FdecDeudaCapital <= ldecPorAplicar,
                        lobjItemFac.FdecDeudaCapital, ldecPorAplicar)
            End If
            If lblnAplicar AndAlso ldecAAplicar > 0 Then
                SGenereItemsRecMedPagoCap(ldecAAplicar, aobjFactura.ObjPrefijo_FactStr.ObjValorPro,
                    aobjFactura.ObjIdFacturaEnt.ObjValorPro,
                    lobjItemFac.ObjIdItemFacturaShr.ObjValorPro)
            End If
            ldecPorAplicar -= ldecAAplicar
            If ldecPorAplicar = 0 Then Exit For
        Next
    End Sub

    ''' <summary>
    ''' Genera un Item de Recibo de Caja por cada medio de pago.
    ''' </summary>
    ''' <param name="aobjFactura">Factura a la cual se le aplicara el pago</param>
    ''' <param name="adecValor">Valor que se le aplicara a la factura y que sera distribuido entre los medios
    ''' de pago generando un item de recibo de caja por cada medio de pago del recibo.</param>
    ''' <param name="ablnAMora">Indica si el valor sera aplicado como un credito a intereses de mora o
    ''' a capital.</param>
    ''' <remarks></remarks>
    Private Sub SGenereItemsRecMedPagoMora(aobjFactura As ClsFactura, adecValor As Decimal)
        Dim lstrCtaDb = String.Empty
        Dim ldecVlrAAplicar = 0D
        Dim ldecTotalAplicado = 0D
        If IsNothing(McolItemsRecCaja) Then
            McolItemsRecCaja = ColItemsRecCaja
        End If
        For Each lobjMedPago As ClsMedioPago In McolMediosPago
            lstrCtaDb = lobjMedPago.ObjIdCtaContabIngresoStr.ObjValorPro
            ldecVlrAAplicar = Math.Round(adecValor * lobjMedPago.FdblTasaParticipaEnPago, 0)
            ldecTotalAplicado += ldecVlrAAplicar
            If lobjMedPago.ObjOrdinal_MedPagoShr.ObjValorPro = McolMediosPago.Count Then
                Dim ldecDif As Decimal = adecValor - ldecTotalAplicado
                If ldecDif <> 0 Then
                    ldecVlrAAplicar += ldecDif
                    ldecTotalAplicado += ldecDif
                End If
            End If
            Dim lobjNewItemRecCaja As ClsItemRecCaja = FobjNewItemRecCaja()
            With lobjNewItemRecCaja
                .ObjIdFactura_ItemRecEnt.ObjValorPro = aobjFactura.ObjIdFacturaEnt.ObjValorPro
                .ObjIdItemFac_ItemRecShr.ObjValorPro = 0
                .ObjIdCuentaDb_ItemRecStr.ObjValorPro = lstrCtaDb
                .ObjIdItemRecCajaShr.ObjValorPro = McolItemsRecCaja.Count + 1
                .ObjIdTipoItemRecByt.ObjValorPro = EnuTipoItemRecCajaDef.EnuAbonoIntMora
                .ObjPrefijoFact_ItemRecStr.ObjValorPro = aobjFactura.ObjPrefijo_FactStr.ObjValorPro
                .ObjBaseDsctoDec.ObjValorPro = 0
                .ObjTasaDsctoDbl.ObjValorPro = 0
                .ObjValor_ItemRecDec.ObjValorPro = ldecVlrAAplicar
            End With
            McolItemsRecCaja.Add(lobjNewItemRecCaja)
        Next
        If adecValor <> ldecTotalAplicado Then
            Throw New ErrorInesperadoPanLException("Total aplicado difiere del valor a aplicar!")
        End If
    End Sub

    Private Sub SGenereItemsRecMedPagoCap(adecValor As Decimal, astrPrefFac As String,
        aentIdFactura As Integer, ashrIdItemFac As Short)
        Dim lstrCtaDb = String.Empty
        Dim ldecVlrAAplicar = 0D
        Dim ldecTotalAplicado = 0D
        For Each lobjMedPago As ClsMedioPago In ColMediosPago
            lstrCtaDb = lobjMedPago.ObjIdCtaContabIngresoStr.ObjValorPro
            ldecVlrAAplicar = Math.Round(adecValor * lobjMedPago.FdblTasaParticipaEnPago, 0)
            ldecTotalAplicado += ldecVlrAAplicar
            If lobjMedPago.ObjOrdinal_MedPagoShr.ObjValorPro = McolMediosPago.Count Then
                Dim ldecDif As Decimal = adecValor - ldecTotalAplicado
                If ldecDif <> 0 Then
                    ldecVlrAAplicar += ldecDif
                    ldecTotalAplicado += ldecDif
                End If
            End If
            Dim lobjNewItemRecCaja As ClsItemRecCaja = FobjNewItemRecCaja()
            With lobjNewItemRecCaja
                .ObjPrefijoFact_ItemRecStr.ObjValorPro = astrPrefFac
                .ObjIdFactura_ItemRecEnt.ObjValorPro = aentIdFactura
                .ObjIdItemFac_ItemRecShr.ObjValorPro = ashrIdItemFac
                .ObjIdCuentaDb_ItemRecStr.ObjValorPro = lstrCtaDb
                .ObjIdItemRecCajaShr.ObjValorPro = McolItemsRecCaja.Count + 1
                .ObjIdTipoItemRecByt.ObjValorPro = EnuTipoItemRecCajaDef.EnuAbonoCapital
                .ObjBaseDsctoDec.ObjValorPro = 0
                .ObjTasaDsctoDbl.ObjValorPro = 0
                .ObjValor_ItemRecDec.ObjValorPro = ldecVlrAAplicar
            End With
            McolItemsRecCaja.Add(lobjNewItemRecCaja)
        Next
        If adecValor <> ldecTotalAplicado Then
            Throw New ErrorInesperadoPanLException("Total aplicado difiere del valor a aplicar!")
        End If
    End Sub

    Private Sub SGenereItemsRecAnticipo()
        If ObjValorAnticipoDec.ObjValorPro > 0 Then
            Dim lstrCtaDb = String.Empty
            Dim ldecVlrAAplicar = 0D
            Dim ldecTotalAplicado = 0D
            Dim ldecVlrAnticipo As Decimal = ObjValorAnticipoDec.ObjValorPro
            If IsNothing(McolItemsRecCaja) Then
                McolItemsRecCaja = ColItemsRecCaja
            End If
            For Each lobjMedPago As ClsMedioPago In McolMediosPago
                lstrCtaDb = lobjMedPago.ObjIdCtaContabIngresoStr.ObjValorPro
                ldecVlrAAplicar = Math.Round(ldecVlrAnticipo * lobjMedPago.FdblTasaParticipaEnPago, 0)
                ldecTotalAplicado += ldecVlrAAplicar
                If lobjMedPago.ObjOrdinal_MedPagoShr.ObjValorPro = McolMediosPago.Count Then
                    Dim ldecDif As Decimal = ldecVlrAnticipo - ldecTotalAplicado
                    If ldecDif <> 0 Then
                        ldecVlrAAplicar += ldecDif
                        ldecTotalAplicado += ldecDif
                    End If
                End If
                Dim lobjNewItemRecCaja As ClsItemRecCaja = FobjNewItemRecCaja()
                With lobjNewItemRecCaja
                    .ObjIdTipoItemRecByt.ObjValorPro = EnuTipoItemRecCajaDef.EnuAnticipo
                    .ObjIdFactura_ItemRecEnt.ObjValorPro = 0
                    .ObjIdItemFac_ItemRecShr.ObjValorPro = 0
                    .ObjIdCuentaDb_ItemRecStr.ObjValorPro = lstrCtaDb
                    .ObjIdItemRecCajaShr.ObjValorPro = McolItemsRecCaja.Count + 1
                    .ObjPrefijoFact_ItemRecStr.ObjValorPro = String.Empty
                    .ObjBaseDsctoDec.ObjValorPro = 0
                    .ObjTasaDsctoDbl.ObjValorPro = 0
                    .ObjValor_ItemRecDec.ObjValorPro = ldecVlrAAplicar
                End With
                McolItemsRecCaja.Add(lobjNewItemRecCaja)
            Next
        End If
    End Sub
#End Region
#Region "Genera Notas Cr, una por Factura"
    Private Sub SGenereNotasCR()
        If Not IsNothing(DtbDescuentos) Then
            Dim larlFrasConDscto = FarlFrasConDscto(), lstrNroFac As String
            Dim lstrNroRecCaj = ClsPanorama.FstrNumeroDcto(ObjPrefijo_RecStr.ObjValorPro,
                ObjIdRecCajaEnt.ObjValorPro)
            For i = 0 To larlFrasConDscto.Count - 1
                lstrNroFac = larlFrasConDscto(i)
                SGenereNotaCrFac(lstrNroFac, lstrNroRecCaj, i = larlFrasConDscto.Count - 1)
            Next
        End If
    End Sub
    Private Function FarlFrasConDscto() As ArrayList
        Dim larlFrasConDscto As New ArrayList, lstrNroFac As String
        For Each ldrwDscto As DataRow In DtbDescuentos.Rows
            lstrNroFac = ldrwDscto("NroFactura")
            If Not larlFrasConDscto.Contains(lstrNroFac) Then
                larlFrasConDscto.Add(lstrNroFac)
            End If
        Next
        larlFrasConDscto.Sort()
        Return larlFrasConDscto
    End Function
    Private Sub SGenereNotaCrFac(astrNroFac As String, astrNroRecCaj As String, ablnUltimo As Boolean)
        Dim lobjFactura As ClsFactura = McolFrasVivas(astrNroFac)
        Dim lstrPreAgr As String = lobjFactura.ObjIdPredioAgrupador_FacStr.ToString
        Dim lstrPref = GobjParametros.FstrPrefijoDoc(EnuIdDocumentoDef.EnuNotaCr)
        Dim lobjNotaCr As New ClsNotaCr(lstrPref)
        lobjNotaCr.SCreeObj(Nothing)
        With lobjNotaCr
            .ObjComentario_NotaCrStr.ObjValorPro = "Nota Crédito generada desde el Recibo de Caja " &
                " Nro " & astrNroRecCaj
            .ObjIdCliente_NotaCrDbl.ObjValorPro = ObjIdCliente_RecDbl.ObjValorPro
            .ObjIdPredioAgrupador_NotaCrStr.ObjValorPro = lstrPreAgr
            .ObjIdTipoNotaCrByt.ObjValorPro = EnuTipoNotaCrDef.EnuDescuento
            .ObjFecha_NotaCrDtm.ObjValorPro = ObjFechaRecDtm.ObjValorPro
            .ObjIdModoNotaCrByt.ObjValorPro = EnuModoNotaCr.EnuPorFactura
            .BlnGenerandoRec = True
        End With
        SAdicioneItemsNcr(lobjNotaCr, astrNroFac)
        lobjNotaCr.SActualice(True)
        SGenereNovsNotaCr(lobjNotaCr, lobjFactura)
        lobjFactura.SActualice(True)
        ObjIdNotasCrStr.ObjValorPro &= lobjNotaCr.StrNumeroNotaCr
        If Not ablnUltimo Then
            ObjIdNotasCrStr.ObjValorPro &= ","
        End If
    End Sub
    Private Sub SAdicioneItemsNcr(aobjNotaCr As ClsNotaCr, astrNroFac As String)
        Dim lstrNroFac As String, lstrPrefFac As String, lentIdFac As Integer
        Dim lstrItemFac As String, lshrIdItemFac As Short
        Dim ldecVlrItem As Decimal, ldecBase As Decimal
        Dim ldblTasa As Double, lentIdTipoDscto As Integer
        Dim lobjFac As New ClsFactura()
        For Each ldrwDscto As DataRow In DtbDescuentos.Rows
            lstrNroFac = ldrwDscto("NroFactura")
            If lstrNroFac = astrNroFac Then
                lstrPrefFac = ClsPanorama.FstrPrefijoDcto(lstrNroFac)
                lentIdFac = ClsPanorama.FentIdDcto(lstrNroFac)
                lobjFac.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefFac, lentIdFac})
                lstrItemFac = ClsPanorama.FobjValorCampo(ldrwDscto("IdItemFact"),
                        EnuTipoValor.EnuString)
                lshrIdItemFac = lstrItemFac.Substring(0, lstrItemFac.IndexOf("-"))
                lentIdTipoDscto = ClsPanorama.FobjValorCampo(ldrwDscto("IdTipoDcto"),
                        EnuTipoValor.EnuInteger)
                ldecBase = ClsPanorama.FobjValorCampo(ldrwDscto("Base"),
                        EnuTipoValor.EnuDecimal)
                ldblTasa = ClsPanorama.FobjValorCampo(ldrwDscto("Tasa"),
                        EnuTipoValor.EnuDouble)
                ldecVlrItem = ClsPanorama.FobjValorCampo(ldrwDscto("Valor"),
                        EnuTipoValor.EnuDecimal)
                Dim lobjNewItemNCr = aobjNotaCr.FobjNewItemNotaCr
                lobjNewItemNCr.ObjBaseDscto_NotaCrDec.ObjValorPro = ldecBase
                lobjNewItemNCr.ObjPrefijoFact_ItemNotaCrStr.ObjValorPro = lstrPrefFac
                lobjNewItemNCr.ObjIdFactura_ItemNotaCrEnt.ObjValorPro = lentIdFac
                lobjNewItemNCr.ObjIdItemFac_ItemNotaCrShr.ObjValorPro = lshrIdItemFac
                lobjNewItemNCr.ObjIdItemNotaCrShr.ObjValorPro = aobjNotaCr.ColItemsNotaCr.Count + 1
                lobjNewItemNCr.ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro = lentIdTipoDscto
                lobjNewItemNCr.ObjTasaDscto_ItemNotaCrDbl.ObjValorPro = ldblTasa
                lobjNewItemNCr.ObjValor_ItemNotaCrDec.ObjValorPro = ldecVlrItem
                aobjNotaCr.SAdicioneNuevoItem(lobjNewItemNCr, lobjFac)
            End If
        Next
    End Sub
    Private Sub SGenereNCrDsctosMora()
        Dim lstrNroFact As String, lobjValorLLave As Object()
        Dim lobjFac As New ClsFactura(), lobjItemFac As ClsItemFactura
        Dim lstrPrefijo As String, lentIdFactura As Integer, lstrIdItemFac As String
        Dim lstrFiltro = "IdTipoDscto = " & EnuTipoDescuento.EnuDsctoIntMora.ToString()
        Dim ldrwDsctosMora As DataRow() = DtbDescuentos.Select(lstrFiltro)
        For Each ldrwDsctoMora As DataRow In ldrwDsctosMora
            lstrNroFact = ClsPanorama.FobjValorCampo(ldrwDsctosMora("NroFactura"),
                    EnuTipoValor.EnuString)
            lstrPrefijo = ClsPanorama.FstrPrefijoDcto(lstrNroFact)
            lstrIdItemFac = ClsPanorama.FobjValorCampo(ldrwDsctosMora("IdItemFact"),
                    EnuTipoValor.EnuString)
            lentIdFactura = ClsPanorama.FentIdDcto(lstrNroFact)
            lobjValorLLave = {GshrIdCarpeta, GshrIdCentroUtil, lstrPrefijo, lentIdFactura}
            lobjFac.SAbra(lobjValorLLave)
            lobjItemFac = lobjFac.ColItemsFactura(lstrIdItemFac)
            If lobjItemFac.ObjTarifaIva_ItemFactDbl.ObjValorPro > 0 Then

            Else

            End If

        Next
    End Sub
    Private Sub SGenereNotaCr(astrNroFac As String, ashrIdItemFac As Short, astrNroRecCaj As String,
            ablnUltimo As Boolean)
        Dim lobjFactura As ClsFactura = McolFrasVivas(astrNroFac)
        Dim lstrPreAgr As String = lobjFactura.ObjIdPredioAgrupador_FacStr.ToString
        Dim lstrPref = GobjParametros.FstrPrefijoDoc(EnuIdDocumentoDef.EnuNotaCr)
        Dim lobjNotaCr As New ClsNotaCr(lstrPref)
        lobjNotaCr.SCreeObj(Nothing)
        With lobjNotaCr
            .ObjComentario_NotaCrStr.ObjValorPro = "Nota Crédito generada desde el Recibo de Caja " &
                " Nro " & astrNroRecCaj
            .ObjIdCliente_NotaCrDbl.ObjValorPro = ObjIdCliente_RecDbl.ObjValorPro
            .ObjIdPredioAgrupador_NotaCrStr.ObjValorPro = lstrPreAgr
            .ObjIdTipoNotaCrByt.ObjValorPro = EnuTipoNotaCrDef.EnuDescuento
            .ObjFecha_NotaCrDtm.ObjValorPro = ObjFechaRecDtm.ObjValorPro
            .BlnGenerandoRec = True
        End With
        SAdicioneItemsNcr(lobjNotaCr, astrNroFac)
        lobjNotaCr.SActualice(True)
        SGenereNovsNotaCr(lobjNotaCr, lobjFactura)
        lobjFactura.SActualice(True)
        ObjIdNotasCrStr.ObjValorPro &= lobjNotaCr.StrNumeroNotaCr
        If Not ablnUltimo Then
            ObjIdNotasCrStr.ObjValorPro &= ","
        End If
    End Sub
    Private Sub SGenereNovsNotaCr(aobjNotaCr As ClsNotaCr, aobjFact As ClsFactura)
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
            For Each lobjItemNotaCr As ClsItemNotaCr In aobjNotaCr.ColItemsNotaCr
                If Not IsNothing(aobjFact) Then
                    If aobjFact.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
                        aobjFact.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
                    End If
                    aobjFact.SApliqueNotaCr(lobjItemNotaCr)
                End If
            Next
        End If
    End Sub
#End Region
#Region "Aplica los Items del Recibo de Caja"
    Private Sub SAplicaItemsRecCaja()
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
            For Each lobjItemRec As ClsItemRecCaja In McolItemsRecCaja
                Dim lenuTipoItemRec As EnuTipoItemRecCajaDef = lobjItemRec.ObjIdTipoItemRecByt.ObjValorPro
                Select Case lenuTipoItemRec
                    Case EnuTipoItemRecCajaDef.EnuAbonoCapital, EnuTipoItemRecCajaDef.EnuAbonoIntMora
                        SApliqueItemPago(lobjItemRec)
                    Case EnuTipoItemRecCajaDef.EnuAnticipo
                        SApliqueItemAnticipo(lobjItemRec)
                    Case EnuTipoItemRecCajaDef.None
                        Throw New ErrorInesperadoPanLException("Tipo Item Rec no valido!")
                End Select
            Next
        End If
    End Sub
    Private Sub SApliqueItemPago(aobjItemRec As ClsItemRecCaja)
        With aobjItemRec
            Dim lstrNroFact As String = ClsPanorama.FstrNumeroDcto(.ObjPrefijoFact_ItemRecStr.ObjValorPro,
                .ObjIdFactura_ItemRecEnt.ObjValorPro)
            Dim lobjFactura As ClsFactura = McolFrasVivas(lstrNroFact)
            lobjFactura.SApliqueCreditoRC(aobjItemRec)
        End With
    End Sub
    Private Sub SApliqueItemAnticipo(aobjItemRec As ClsItemRecCaja)
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
            Dim lstrIdPredioAgr = FstrIdPpredioAgrAnticipo()
            If IsNothing(MobjAnticipo) Then
                MobjAnticipo = MobjClienteRec.FobjNuevoAnticipo(ObjFechaRecDtm.ObjValorPro,
                    lstrIdPredioAgr)
                With MobjAnticipo
                    .ObjIdTipoDocOrigen_AntByt.ObjValorPro = EnuTipoDocOri.EnuReciboCaja
                    .ObjServicios_AntStr.ObjValorPro = ObjServicios_RecStr.ObjValorPro
                    .ObjIdPredio_AntStr.ObjValorPro = String.Empty
                    .ObjIdDocOrigen_AntEnt.ObjValorPro = ObjIdRecCajaEnt.ObjValorPro
                    .ObjPrefijoDocOrigen_AntStr.ObjValorPro = ObjPrefijo_RecStr.ObjValorPro
                End With
            End If
            With aobjItemRec
                MobjAnticipo.SGenereNovedadAntRecibido(ObjFechaRecDtm.ObjValorPro,
                    .ObjIdCuentaDb_ItemRecStr.ObjValorPro, .ObjValor_ItemRecDec.ObjValorPro)
            End With
        End If
    End Sub
    ''' <summary>
    ''' Devuelve el Id del Predio Agrupador al cual se le va a aplicar el anticipo.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>Solo puede haber un Anticipo por recibo de caja. Si hay varios predios 
    ''' agrupadores en el recibo de caja se le aplica al primer predio del arreglo diferente
    ''' a "Sin Predio Agrupador"</remarks>
    Private Function FstrIdPpredioAgrAnticipo() As String
        Dim lstrPredAgrupadores = ObjIdPredioAgrupador_RecStr.ToString.Split(",")
        Dim lstrIdPrediosAgr As ArrayList = ClsOrionCop.FstrIdPrediosAgr(lstrPredAgrupadores)
        Dim lstrIdPredioAgr = String.Empty
        For Each lstrIdPredio As String In lstrIdPrediosAgr
            If lstrIdPredio <> "" Then
                lstrIdPredioAgr = lstrIdPredio
                Exit For
            End If
        Next
        Return lstrIdPredioAgr
    End Function
#End Region
    ''' <summary>
    ''' Devuelve la coleccion de facturas vivas al momento de generar el recibo de caja
    ''' correspondientes al cliente y al predio agrupador especificados.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function FcolFrasVivas() As Collection 'Ok
        Static lstrIdPredAgr As String() = Nothing
        Dim lblnConstruir As Boolean = McolFrasVivas Is Nothing
        If Not lblnConstruir Then
            lblnConstruir = Not FblnStrArrayIguales(ObjIdPredioAgrupador_RecStr.ToString.Split(","),
                lstrIdPredAgr)
        End If
        If lblnConstruir Then
            lstrIdPredAgr = ObjIdPredioAgrupador_RecStr.ToString.Split(",")
            McolFrasVivas = MobjClienteRec.FcolFacturas(lstrIdPredAgr, {"A"}, True)
        End If
        Return McolFrasVivas
    End Function
    Friend ReadOnly Property ColItemsRecCaja As Collection
        Get
            If IsNothing(McolItemsRecCaja) Then
                McolItemsRecCaja = New Collection
                SCargueDtbItemsRecCaja()
                If ObjIdRecCajaEnt.BlnEsValido AndAlso EnuEstadoActualizacion <> EnuEstadoObjetoDef.EnuCreando Then
                    If Not IsNothing(MdtbItemsRecCaja) AndAlso MdtbItemsRecCaja.Rows.Count > 0 Then
                        Dim ldrwItemsRec() As DataRow = MdtbItemsRecCaja.Select
                        For Each ldrwItemRecCaj As DataRow In ldrwItemsRec
                            Dim lobjItemRec As New ClsItemRecCaja(Me, ldrwItemRecCaj)
                            lobjItemRec.SLeaValores(True)
                            McolItemsRecCaja.Add(lobjItemRec, lobjItemRec.ObjIdItemRecCajaShr.ToString)
                        Next
                    End If
                End If
            End If
            Return McolItemsRecCaja
        End Get
    End Property
    ''' <summary>
    ''' Devuelve un nuevo item de recibo de caja en modo de crear y con las propiedades generales asignadas
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function FobjNewItemRecCaja() As ClsItemRecCaja
        Dim ldrwNewItemRecCaja As DataRow = MdtbItemsRecCaja.NewRow
        Dim lobjNewItemRecCaja As New ClsItemRecCaja(Me, ldrwNewItemRecCaja)
        With lobjNewItemRecCaja
            .SCreeObj(Nothing)
            .ObjIdCarpeta_ItemRecShr.ObjValorPro = GshrIdCarpeta
            .ObjIdCentroUtil_ItemRecShr.ObjValorPro = GshrIdCentroUtil
            .ObjIdRecCaja_ItemRecEnt.ObjValorPro = ObjIdRecCajaEnt.ObjValorPro
            .ObjPrefijoRec_ItemRecStr.ObjValorPro = ObjPrefijo_RecStr.ObjValorPro
        End With
        Return lobjNewItemRecCaja
    End Function
    Private Sub SCargueDtbItemsRecCaja()
        If IsNothing(MdtbItemsRecCaja) Then
            Dim lstrIdRecCaja = ObjIdRecCajaEnt.ToString
            If lstrIdRecCaja = String.Empty Then lstrIdRecCaja = "0"
            Dim lstrIndice = {{ClsPrefijoFact_ItemRecStr.SstrNombreCampoBd, "ASC"},
                            {ClsIdFactura_ItemRecEnt.SstrNombreCampoBd, "ASC"},
                            {ClsIdItemFac_ItemRecShr.SstrNombreCampoBd, "ASC"}}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsPrefijo_RecStr.SstrNombreCampoBd &
                " = '" & ObjPrefijo_RecStr.ToString & "' AND " & ClsIdRecCaja_ItemRecEnt.SstrNombreCampoBd &
                " = " & lstrIdRecCaja
            Dim lstrCamposSelect() = {"*"}
            MdtbItemsRecCaja = ClsPanorama.FdtbDataTable(ClsItemRecCaja.SstrNombreTabla, lstrCamposSelect,
                lstrIndice, lstrFiltro, False, Array.Empty(Of String))
        End If
    End Sub
    Friend Function FdtbInfItemsRC()
        Const lstrFacMora = "Todas"
        Dim lstrPreRec = String.Empty, lentIdRec = 0
        Dim lstrPreFac = String.Empty, lentIdFac = 0, lshrIdItemFac = 0S
        Dim lstrFac = String.Empty
        Dim ldecValor As Decimal(), lstrDetalle As String()
        Dim ldtbInfItemsRC = FdtbInfItems()
        Dim ldrwNew As DataRow = Nothing
        ReDim ldecValor(8)
        ReDim lstrDetalle(8)
        For i = 0 To 7
            ldecValor(i) = -1
        Next
        Dim ldtbItemsRC = FdtbInfItemsRecCaja(ObjPrefijo_RecStr.ObjValorPro,
                ObjIdRecCajaEnt.ObjValorPro, ObjIdRecCajaEnt.ObjValorPro)
        SPuebleInfPagoCap(ldtbInfItemsRC, ldtbItemsRC)
        For Each ldrwItemRec As DataRow In ldtbItemsRC.Rows
            Dim lenuTipoItemRec As EnuTipoItemRecCajaDef = ClsPanorama.FobjValorCampo(
                    ldrwItemRec(ClsIdTipoItemRecByt.SstrNombreCampoBd),
                    EnuTipoValor.EnuInteger)
            lstrPreRec = ClsPanorama.FobjValorCampo(ldrwItemRec(
                    ClsPrefijo_RecStr.SstrNombreCampoBd), EnuTipoValor.EnuString)
            lentIdRec = ClsPanorama.FobjValorCampo(ldrwItemRec(
                    ClsIdRecCajaEnt.SstrNombreCampoBd), EnuTipoValor.EnuInteger)
            lstrPreFac = ClsPanorama.FobjValorCampo(ldrwItemRec(
                    ClsPrefijoFact_ItemRecStr.SstrNombreCampoBd),
                    EnuTipoValor.EnuString)
            lentIdFac = ClsPanorama.FobjValorCampo(ldrwItemRec(
                    ClsIdFactura_ItemRecEnt.SstrNombreCampoBd),
                    EnuTipoValor.EnuInteger)
            lshrIdItemFac = ClsPanorama.FobjValorCampo(ldrwItemRec(
                    ClsIdItemFac_ItemRecShr.SstrNombreCampoBd),
                    EnuTipoValor.EnuShort)
            lstrFac = ClsPanorama.FstrNumeroDcto(lstrPreFac, lentIdFac)
            Select Case lenuTipoItemRec
                Case EnuTipoItemRecCajaDef.EnuAbonoCapital
                '
                Case EnuTipoItemRecCajaDef.EnuAbonoIntMora
                    lstrDetalle(0) = My.Resources.RCAbonoMor
                    ldecValor(0) = 0
                    ldecValor(0) += ClsPanorama.FobjValorCampo(ldrwItemRec(ClsValor_ItemRecDec.SstrNombreCampoBd),
                            EnuTipoValor.EnuDecimal)
                Case EnuTipoItemRecCajaDef.EnuAnticipo
                    lstrDetalle(1) = My.Resources.RCAnticipo & FstrComplementoSerAnticipo()
                    ldecValor(1) = 0
                    ldecValor(1) += ClsPanorama.FobjValorCampo(ldrwItemRec(ClsValor_ItemRecDec.SstrNombreCampoBd),
                            EnuTipoValor.EnuDecimal)
                Case EnuTipoItemRecCajaDef.EnuDsctoCapital
                    lstrDetalle(2) = My.Resources.RCDsctoCap
                    ldecValor(2) = 0
                    ldecValor(2) += ClsPanorama.FobjValorCampo(ldrwItemRec(ClsValor_ItemRecDec.SstrNombreCampoBd),
                            EnuTipoValor.EnuDecimal)
                Case EnuTipoItemRecCajaDef.EnuDsctoPP
                    lstrDetalle(3) = My.Resources.RCDsctoPP
                    ldecValor(3) = 0
                    ldecValor(3) += ClsPanorama.FobjValorCampo(ldrwItemRec(ClsValor_ItemRecDec.SstrNombreCampoBd),
                    EnuTipoValor.EnuDecimal)
                Case EnuTipoItemRecCajaDef.EnuDsctoIntMora
                    lstrDetalle(4) = My.Resources.RCDsctoMor
                    ldecValor(4) = 0
                    ldecValor(4) += ClsPanorama.FobjValorCampo(ldrwItemRec(ClsValor_ItemRecDec.SstrNombreCampoBd),
                    EnuTipoValor.EnuDecimal)
                Case EnuTipoItemRecCajaDef.EnuReteFuente
                    lstrDetalle(5) = My.Resources.RCReteFte
                    ldecValor(5) = 0
                    ldecValor(5) += ClsPanorama.FobjValorCampo(ldrwItemRec(ClsValor_ItemRecDec.SstrNombreCampoBd),
                    EnuTipoValor.EnuDecimal)
                Case EnuTipoItemRecCajaDef.EnuReteIca
                    lstrDetalle(6) = My.Resources.RCRetIca
                    ldecValor(6) = 0
                    ldecValor(6) += ClsPanorama.FobjValorCampo(ldrwItemRec(ClsValor_ItemRecDec.SstrNombreCampoBd),
                    EnuTipoValor.EnuDecimal)
                Case EnuTipoItemRecCajaDef.EnuReteIva
                    lstrDetalle(7) = My.Resources.RCRetIva
                    ldecValor(7) = 0
                    ldecValor(7) += ClsPanorama.FobjValorCampo(ldrwItemRec(ClsValor_ItemRecDec.SstrNombreCampoBd),
                    EnuTipoValor.EnuDecimal)
            End Select
        Next
        For i = 0 To 7
            If ldecValor(i) >= 0 Then
                ldrwNew = ldtbInfItemsRC.NewRow
                ldrwNew("Ordinal") = ldtbInfItemsRC.Rows.Count + 1
                ldrwNew("Prefijo") = ObjPrefijo_RecStr.ObjValorPro
                ldrwNew("IdReciboCaja") = ObjIdRecCajaEnt.ObjValorPro
                If i = 0 Then
                    ldrwNew("NroFact") = lstrFacMora
                ElseIf i <> 1 Then
                    ldrwNew("NroFact") = lstrFac
                Else
                    ldrwNew("NroFact") = "0"
                End If
                ldrwNew("Detalle") = lstrDetalle(i)
                ldrwNew("Valor") = ldecValor(i)
                ldtbInfItemsRC.Rows.Add(ldrwNew)
                ldecValor(i) = 0
                lstrDetalle(i) = ""
            End If
        Next
        Return ldtbInfItemsRC
    End Function
    Friend Shared Function FdtbInfItemsRecCaja(astrPrefRC As String,
            aentIdRecCajIni As Integer, aentIdRecCajaFin As Integer) As DataTable
        Dim lstrCamposSelect = {ClsPrefijo_RecStr.SstrNombreCampoBd,
                ClsIdRecCajaEnt.SstrNombreCampoBd,
                ClsIdItemRecCajaShr.SstrNombreCampoBd,
                ClsPrefijoFact_ItemRecStr.SstrNombreCampoBd,
                ClsIdFactura_ItemRecEnt.SstrNombreCampoBd,
                ClsIdItemFac_ItemRecShr.SstrNombreCampoBd,
                ClsIdTipoItemRecByt.SstrNombreCampoBd,
                "SUM(" & ClsValor_ItemRecDec.SstrNombreCampoBd & ") AS Valor"}
        Dim lstrCamposGroup = {ClsPrefijo_RecStr.SstrNombreCampoBd,
                ClsIdRecCajaEnt.SstrNombreCampoBd,
                ClsPrefijoFact_ItemRecStr.SstrNombreCampoBd,
                ClsIdFactura_ItemRecEnt.SstrNombreCampoBd,
                ClsIdItemFac_ItemRecShr.SstrNombreCampoBd,
                ClsIdTipoItemRecByt.SstrNombreCampoBd}
        Dim lstrIndice = {{ClsIdItemRecCajaShr.SstrNombreCampoBd, "ASC"}}
        Dim lstrCamposAgr = {}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsPrefijo_RecStr.SstrNombreCampoBd &
                " = '" & astrPrefRC & "' AND " &
                ClsIdRecCaja_ItemRecEnt.SstrNombreCampoBd & " BETWEEN " &
                aentIdRecCajIni & " AND " & aentIdRecCajaFin
        Dim ldtbInfItemsRC = ClsPanorama.FdtbDataTable(ClsItemRecCaja.SstrNombreTabla,
                lstrCamposSelect, lstrIndice, lstrFiltro, False, lstrCamposGroup)
        Return ldtbInfItemsRC
    End Function
    Friend Shared Function FdtbInfItems() As DataTable
        Dim ldtbInfItemsRC = New DataTable
        Dim ldclOrdinal = New DataColumn("Ordinal", System.Type.GetType("System.Int32"))
        Dim ldclPrefijo = New DataColumn("Prefijo", System.Type.GetType("System.String"))
        Dim ldclIdRecCaja = New DataColumn("IdReciboCaja", System.Type.GetType("System.Int32"))
        Dim ldclNroFac = New DataColumn("NroFact", System.Type.GetType("System.String"))
        Dim ldclDet = New DataColumn("Detalle", System.Type.GetType("System.String"))
        Dim ldclValor = New DataColumn("Valor", System.Type.GetType("System.Decimal"))
        ldtbInfItemsRC.Columns.Add(ldclOrdinal)
        ldtbInfItemsRC.Columns.Add(ldclPrefijo)
        ldtbInfItemsRC.Columns.Add(ldclIdRecCaja)
        ldtbInfItemsRC.Columns.Add(ldclNroFac)
        ldtbInfItemsRC.Columns.Add(ldclDet)
        ldtbInfItemsRC.Columns.Add(ldclValor)
        Dim lcolPKIndice(2) As DataColumn
        lcolPKIndice(0) = ldclPrefijo
        lcolPKIndice(1) = ldclIdRecCaja
        lcolPKIndice(2) = ldclOrdinal
        ldtbInfItemsRC.PrimaryKey = lcolPKIndice
        Return ldtbInfItemsRC
    End Function
    Private Sub SPuebleInfPagoCap(adtbInfItemsRC As DataTable, adtbItemsRC As DataTable)
        Dim lstrPre = String.Empty, lentId = 0, lshrIdItem = 0S, i = 0
        Dim lstrPreFac = String.Empty, lentIdFac = 0, lshrIdItemFac = 0S
        Dim lstrNroFac = String.Empty, lstrPreRec = String.Empty, lentIdRec = 0
        Dim lstrNroRec = String.Empty, lstrNroRecAct = String.Empty
        Dim lstrDetItem = String.Empty, ldecValor = 0D
        Dim lobjFact = New ClsFactura()
        Dim lobjItemFac As ClsItemFactura = Nothing
        Dim lobjValorLlave As Object() = Nothing
        Dim lenuTipoItemRec As EnuTipoItemRecCajaDef = EnuTipoItemRecCajaDef.None
        For Each ldrwItemRec As DataRow In adtbItemsRC.Rows
            lenuTipoItemRec = ClsPanorama.FobjValorCampo(ldrwItemRec("IdTipoItemRec"),
                    EnuTipoValor.EnuInteger)
            If lenuTipoItemRec = EnuTipoItemRecCajaDef.EnuAbonoCapital Then
                lstrPreRec = ClsPanorama.FobjValorCampo(ldrwItemRec(
                        ClsPrefijo_RecStr.SstrNombreCampoBd), EnuTipoValor.EnuString)
                lentIdRec = ClsPanorama.FobjValorCampo(ldrwItemRec(
                        ClsIdRecCajaEnt.SstrNombreCampoBd), EnuTipoValor.EnuInteger)
                lstrPreFac = ClsPanorama.FobjValorCampo(ldrwItemRec(
                        ClsPrefijoFact_ItemRecStr.SstrNombreCampoBd),
                        EnuTipoValor.EnuString)
                lentIdFac = ClsPanorama.FobjValorCampo(ldrwItemRec(
                        ClsIdFactura_ItemRecEnt.SstrNombreCampoBd),
                        EnuTipoValor.EnuInteger)
                lshrIdItemFac = ClsPanorama.FobjValorCampo(ldrwItemRec(
                        ClsIdItemFac_ItemRecShr.SstrNombreCampoBd),
                        EnuTipoValor.EnuShort)
                lstrNroRec = ClsPanorama.FstrNumeroDcto(lstrPreRec, lentIdRec)
                If lstrNroRecAct <> lstrNroRec Then
                    lstrNroRecAct = lstrNroRec
                    i = 1
                Else
                    i += 1
                End If
                If lenuTipoItemRec = EnuTipoItemRecCajaDef.EnuAbonoCapital Then
                    lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrPreFac,
                            lentIdFac}
                    lobjFact.SAbra(lobjValorLlave)
                End If
                If lshrIdItemFac > 0 Then
                    lobjItemFac = lobjFact.ColItemsFactura(lshrIdItemFac.ToString)
                    lstrDetItem = My.Resources.RCAbonoCap & " " &
                        lobjItemFac.ObjDetalle_ItemFactStr.ObjValorNuevo
                Else
                    lstrDetItem = "Anticipo recibido"
                End If
                ldecValor = ClsPanorama.FobjValorCampo(ldrwItemRec(
                        ClsValor_ItemRecDec.SstrNombreCampoBd),
                        EnuTipoValor.EnuDecimal)
                Dim ldrwNew As DataRow = adtbInfItemsRC.NewRow
                lstrNroFac = ClsPanorama.FstrNumeroDcto(lstrPreFac, lentIdFac)
                ldrwNew("Ordinal") = i
                ldrwNew("Prefijo") = ObjPrefijo_RecStr.ObjValorPro
                ldrwNew("IdReciboCaja") = ObjIdRecCajaEnt.ObjValorPro
                ldrwNew("NroFact") = lstrNroFac
                ldrwNew("Detalle") = lstrDetItem
                ldrwNew("Valor") = ldecValor
                adtbInfItemsRC.Rows.Add(ldrwNew)
            End If
        Next
    End Sub
    Friend Function FstrComplementoSerAnticipo() As String
        Dim lstrComp = String.Empty
        Dim lstrServiciosAnt As String = ObjAnticipo.ObjServicios_AntStr.ObjValorPro
        If String.IsNullOrEmpty(lstrServiciosAnt) OrElse lstrServiciosAnt.Contains("A") Then
            lstrComp = " - Para aplicar a todos los servicios!"
        ElseIf lstrServiciosAnt.Contains(",") Then
            lstrComp = " - Para aplicar a cualquiera de los servicios incluidos en el Pago!"
        Else
            If lstrServiciosAnt = "0" Then
                lstrComp = " - Para aplicar a la Cuota de Administración!"
            Else
                If IsNumeric(lstrServiciosAnt) Then
                    Dim lstrNomSer = GobjParametros.FstrNombreServicio(lstrServiciosAnt)
                    lstrComp = " - Para aplicar al servicio " & lstrNomSer
                End If
            End If
        End If
        Return lstrComp
    End Function
#End Region

#Region "Manejo medios de Pago"
    Friend ReadOnly Property ColMediosPago As Collection
        Get
            If IsNothing(McolMediosPago) Then
                McolMediosPago = New Collection
            End If
            If ObjIdRecCajaEnt.BlnEsValido AndAlso EnuEstadoActualizacion <>
                    EnuEstadoObjetoDef.EnuCreando AndAlso McolMediosPago.Count = 0 Then
                McolMediosPago.Clear()
                SCargueDtbMediosPago()
                If Not IsNothing(MdtbMediosPago) AndAlso MdtbMediosPago.Rows.Count > 0 Then
                    Dim ldrwMediosPago() As DataRow = MdtbMediosPago.Select
                    For Each ldrwMedPago As DataRow In ldrwMediosPago
                        Dim lobjMedioPago As New ClsMedioPago(Me, ldrwMedPago)
                        lobjMedioPago.SLeaValores(True)
                        McolMediosPago.Add(lobjMedioPago, lobjMedioPago.ObjOrdinal_MedPagoShr.ToString)
                    Next
                End If
            End If
            Return McolMediosPago
        End Get
    End Property
    Friend ReadOnly Property DtbMediosPago As DataTable
        Get
            SCargueDtbMediosPago()
            SComplementeDtbMedPago()
            Return MdtbMediosPago
        End Get
    End Property
    Private Sub SCargueDtbMediosPago()
        If IsNothing(MdtbMediosPago) Then
            Dim lstrIdRecCaja = ObjIdRecCajaEnt.ToString
            If String.IsNullOrEmpty(lstrIdRecCaja) Then lstrIdRecCaja = "0"
            Dim lstrIndice = {{StrCampoCarpeta, "ASC"},
                            {StrCampoCentroUtil, "ASC"},
                            {ClsPrefijo_RecStr.SstrNombreCampoBd, "ASC"},
                            {ClsIdRecCaja_MedPagoEnt.SstrNombreCampoBd, "ASC"},
                            {ClsOrdinal_MedPagoShr.SstrNombreCampoBd, "ASC"}}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsPrefijo_RecStr.SstrNombreCampoBd &
                " = '" & ObjPrefijo_RecStr.ToString & "' AND " & ClsIdRecCaja_ItemRecEnt.SstrNombreCampoBd &
                " = " & lstrIdRecCaja
            Dim lstrCamposSelect() = {"*", "'*' AS NombreMedioPago", "'*' AS NombreCuenta"}
            MdtbMediosPago = ClsPanorama.FdtbDataTable(ClsMedioPago.SstrNombreTabla, lstrCamposSelect, lstrIndice,
                lstrFiltro)
        End If
    End Sub
    Private Sub SComplementeDtbMedPago()
        Dim ldrwMediosPago() As DataRow = MdtbMediosPago.Select
        If ldrwMediosPago.Length > 0 Then
            Dim lbytIdTipoMedPago As Byte
            Dim lstrNombreMedPago As String
            Dim lstrIdCtaCont As String
            Dim lstrNombreCuenta As String
            For Each ldrwMedPag As DataRow In ldrwMediosPago
                lbytIdTipoMedPago = ClsPanorama.FobjValorCampo(ldrwMedPag(
                    ClsIdTipoMedPago_MedPagoByt.SstrNombreCampoBd),
                    EnuTipoValor.EnuByte)
                lstrIdCtaCont = ClsPanorama.FobjValorCampo(ldrwMedPag(
                    ClsIdCtaContabIngresoStr.SstrNombreCampoBd), EnuTipoValor.EnuString)
                lstrNombreMedPago = ClsOrionCop.FstrNombreDatoConstanteOri(EnuGrupoConstantesOriDef.EnuMediosPago,
                    lbytIdTipoMedPago)
                lstrNombreCuenta = ClsOrionCop.FstrCuentaBanco(lstrIdCtaCont)
                ldrwMedPag("NombreMedioPago") = lstrNombreMedPago
                ldrwMedPag("NombreCuenta") = lstrNombreCuenta
            Next
        End If
    End Sub
    Friend Function FobjNewMedioPago() As ClsMedioPago
        SCargueDtbMediosPago()
        Dim ldrwNewMedioPago As DataRow = MdtbMediosPago.NewRow
        Dim lobjMedioPago As New ClsMedioPago(Me, ldrwNewMedioPago)
        With lobjMedioPago
            .SCreeObj(Nothing)
        End With
        Return lobjMedioPago
    End Function
    Friend Sub SAdicioneMedioPago(aobjMedioPago As ClsMedioPago)
        If IsNothing(McolMediosPago) Then
            McolMediosPago = ColMediosPago
        End If
        Dim lshrOrdinal As Short = McolMediosPago.Count + 1
        aobjMedioPago.ObjOrdinal_MedPagoShr.ObjValorPro = lshrOrdinal
        McolMediosPago.Add(aobjMedioPago, lshrOrdinal.ToString)
        SAdicionesDataRowMedPago(aobjMedioPago)
    End Sub
    Friend Function FblnEsUnicoMediPago(astrIdCntaContIng As String, astrNroMediPago As String,
                                       aenuTipoMediPago As EnuTipoMedioPagoDef) As Boolean
        Dim lblnEsUnico = True
        If ColMediosPago.Count > 0 Then
            For Each lobjMedPag As ClsMedioPago In McolMediosPago
                With lobjMedPag
                    If .ObjIdTipoMedPago_MedPagoByt.ObjValorPro = aenuTipoMediPago AndAlso
                            .ObjNumeroMedPagoStr.ObjValorNuevo = astrNroMediPago AndAlso
                            .ObjIdCtaContabIngresoStr.ObjValorNuevo = astrIdCntaContIng Then
                        lblnEsUnico = False
                        Exit For
                    End If
                End With
            Next
        End If
        Return lblnEsUnico
    End Function
    Private Sub SAdicionesDataRowMedPago(aobjMedioPago As ClsMedioPago)
        If Not IsNothing(aobjMedioPago) Then
            Dim ldrwNewMedPago = MdtbMediosPago.NewRow
            With aobjMedioPago
                ldrwNewMedPago(StrCampoCarpeta) = GshrIdCarpeta
                ldrwNewMedPago(StrCampoCentroUtil) = GshrIdCentroUtil
                ldrwNewMedPago(ClsPrefijo_RecStr.SstrNombreCampoBd) = String.Empty
                ldrwNewMedPago(ClsIdRecCaja_MedPagoEnt.SstrNombreCampoBd) = 0
                ldrwNewMedPago(ClsValor_MedPagoDec.SstrNombreCampoBd) = .ObjValor_MedPagoDec.ObjValorPro
                ldrwNewMedPago(ClsIdTipoMedPago_MedPagoByt.SstrNombreCampoBd) = .ObjIdTipoMedPago_MedPagoByt.ObjValorPro
                ldrwNewMedPago(ClsIdCtaContabIngresoStr.SstrNombreCampoBd) = .ObjIdCtaContabIngresoStr.ObjValorPro
                ldrwNewMedPago(ClsOrdinal_MedPagoShr.SstrNombreCampoBd) = .ObjOrdinal_MedPagoShr.ObjValorPro
                ldrwNewMedPago(ClsNumeroMedPagoStr.SstrNombreCampoBd) = .ObjNumeroMedPagoStr.ObjValorPro
            End With
            MdtbMediosPago.Rows.Add(ldrwNewMedPago)
            SComplementeDtbMedPago()
        End If
    End Sub
    Friend Sub SLimpieMediosPago()
        If McolMediosPago IsNot Nothing Then
            For i = 1 To McolMediosPago.Count
                SElimineMedPago(i)
                i += 1
            Next
        End If
    End Sub
    Friend Sub SElimineMedPago(ashrOrdinal As Short)
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando AndAlso McolMediosPago.Count > 0 Then
            Dim lblnEncontrado = False
            Dim i = 0
            For Each lobjMedPago As ClsMedioPago In McolMediosPago
                i += 1
                If lobjMedPago.ObjOrdinal_MedPagoShr.ObjValorPro = ashrOrdinal Then
                    lblnEncontrado = True
                    Exit For
                End If
            Next
            If lblnEncontrado Then
                McolMediosPago.Remove(i)
                If McolMediosPago.Count > 0 Then
                    SReorganiceColMePago()
                End If
                SElimineDrwMedPago(ashrOrdinal)
            Else
                Throw New ErrorInesperadoPanLException("Objeto medio de pago no encontrado!")
            End If
        End If
    End Sub
    Private Sub SReorganiceColMePago()
        Dim lcolNewMediosPago As New Collection
        For Each lobjMedPago As ClsMedioPago In McolMediosPago
            lobjMedPago.ObjOrdinal_MedPagoShr.ObjValorPro = lcolNewMediosPago.Count + 1
            lcolNewMediosPago.Add(lobjMedPago, lobjMedPago.ObjOrdinal_MedPagoShr.ToString)
        Next
        McolMediosPago = lcolNewMediosPago
    End Sub
    Private Sub SElimineDrwMedPago(ashrOrdinal As Short)
        Dim lblnEncontrado = False
        Dim ldrwMedPago As DataRow = Nothing
        For i = 0 To MdtbMediosPago.Rows.Count - 1
            ldrwMedPago = MdtbMediosPago.Rows(i)
            If ldrwMedPago(ClsOrdinal_MedPagoShr.SstrNombreCampoBd) = ashrOrdinal Then
                lblnEncontrado = True
                Exit For
            End If
        Next
        If lblnEncontrado Then
            MdtbMediosPago.Rows.Remove(ldrwMedPago)
            Dim j = 0
            Dim ldrwMediosPago() As DataRow = MdtbMediosPago.Select()
            For Each ldrwMedioPago As DataRow In ldrwMediosPago
                j += 1
                ldrwMedioPago(ClsOrdinal_MedPagoShr.SstrNombreCampoBd) = j
            Next
        Else
            Throw New ErrorInesperadoPanLException("DataRow medio de pago no encontrado!")
        End If
    End Sub
    Friend Function FblnTotalMPValido() As Boolean
        Dim lblnEsValido = (FdecValorTotalMedPago() = ObjValor_RecDec.ObjValorPro)
        If Not lblnEsValido Then
            If FdecValorTotalMedPago() > ObjValor_RecDec.ObjValorPro Then
                Dim lenuSevNot As EnuSeveridadNot = EnuSeveridadNot.EnuDatoInvalido
                Dim lstrMens = "El total de los Medios de Pago es mayor al Valor Recibido!"
                SLevanteEventoNot(lstrMens, "", EnuIdMens.EnuTotalMP, lenuSevNot)
            End If
        End If
        Return lblnEsValido
    End Function
    Friend Function FdecValorTotalMedPago() As Decimal
        Dim ldecVlrTotal = 0D
        If IsNothing(McolMediosPago) Then
            McolMediosPago = ColMediosPago
        End If
        For Each lobjMedPago As ClsMedioPago In McolMediosPago
            ldecVlrTotal += lobjMedPago.ObjValor_MedPagoDec.ObjValorPro
        Next
        Return ldecVlrTotal
    End Function
#End Region

#Region "Anticipo"
    Friend ReadOnly Property ObjAnticipo As ClsAnticipo
        Get
            If IsNothing(MobjAnticipo) Then
                Dim ldrwAnticipo As DataRow = FdrwAnticipo()
                If Not IsNothing(ldrwAnticipo) Then
                    MobjAnticipo = New ClsAnticipo(ObjClienteRecibo, ldrwAnticipo)
                    MobjAnticipo.SLeaValores(True)
                    HblnEsAnulable = True
                End If
            End If
            Return MobjAnticipo
        End Get
    End Property
    Private Function FdrwAnticipo()
        Dim ldrwAnticipo As DataRow = Nothing
        Dim lstrfiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsIdTipoDocOrigen_AntByt.SstrNombreCampoBd &
            " = " & EnuTipoDocOri.EnuReciboCaja & " AND " &
            ClsPrefijoDocOrigen_AntStr.SstrNombreCampoBd & " = '" & ObjPrefijo_RecStr.ObjValorPro &
            "' AND " & ClsIdDocOrigen_AntEnt.SstrNombreCampoBd & " = " & ObjIdRecCajaEnt.ObjValorPro
        Dim ldrwAnticipos = ClsPanorama.FdrwDataRow(ClsAnticipo.SstrNombreTabla, {"*"},
            {{ClsPrefijoDocOrigen_AntStr.SstrNombreCampoBd, "ASC"},
                {ClsIdDocOrigen_AntEnt.SstrNombreCampoBd, "ASC"}}, lstrfiltro)
        If ldrwAnticipos.Length > 0 Then
            ldrwAnticipo = ldrwAnticipos(0)
        End If
        Return ldrwAnticipo
    End Function
#End Region

#Region "Nota RCr"
    Private Sub SGenereNotaRevCr()
        Dim lstrPref = GobjParametros.FstrPrefijoDoc(EnuIdDocumentoDef.EnuNotaReversaCr)
        Dim lobjNotaRcr As New ClsNotaReversionCr(lstrPref)
        lobjNotaRcr.SCreeObj(Nothing)
        With lobjNotaRcr
            Dim ldtmFechaNotaRRC As Date = GCDTMFECHANULA
            Dim ldtmFechaAnu As Date = ObjFechaAnulacion_RecDtm.ObjValorPro
            If IsDate(ldtmFechaAnu) Then
                ldtmFechaNotaRRC = ldtmFechaAnu.Date
            Else
                Throw New ErrorInesperadoPanLException("Se esperaba una fecha!")
            End If
            .ObjTipoDocReversadoByt.ObjValorPro = EnuDocReversado.EnuReciboC
            .ObjFecha_NotaReversaCrDtm.ObjValorPro = ldtmFechaNotaRRC
            .ObjDetalle_NotaReversaCrStr.ObjValorPro = "Anulacion Recibo Caja " & StrNumeroRecCaja
            .ObjIdPredioAgrupador_NotaReversaCrStr.ObjValorPro = ObjIdPredioAgrupador_RecStr.ObjValorPro
            .ObjPrefijoDoc_NotaReversaCrStr.ObjValorPro = ObjPrefijo_RecStr.ObjValorPro
            .ObjIdDoc_NotaReversaCrEnt.ObjValorPro = ObjIdRecCajaEnt.ObjValorPro
            .ObjIdCliente_NotaReversaCrDbl.ObjValorPro = ObjIdCliente_RecDbl.ObjValorPro
            .ObjValor_NotaReversaCrDec.ObjValorPro = ObjValor_RecDec.ObjValorPro
            .ObjDocReversado = Me
            .SActualice(True)
        End With
    End Sub
    Friend ReadOnly Property StrNroNotaRCr As String
        Get
            Dim lentIdRC = 0, lstrPrefRC = String.Empty
            If Not IsNothing(ObjIdRecCajaEnt.ObjValorPro) Then
                lstrPrefRC = ObjPrefijo_RecStr.ObjValorPro
                lentIdRC = ObjIdRecCajaEnt.ObjValorPro
            End If
            If String.IsNullOrEmpty(MstrNroNotaRCr) Then
                Dim lstrCamposSelect As String() = {ClsPrefijo_NotaReversaCrStr.SstrNombreCampoBd,
                        ClsIdNotaReversaCrEnt.SstrNombreCampoBd}
                Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                        ClsTipoDocReversadoByt.SstrNombreCampoBd & " = " &
                        EnuDocReversado.EnuReciboC & " AND " &
                        ClsPrefijoDoc_NotaReversaCrStr.SstrNombreCampoBd & " = '" &
                        lstrPrefRC & "' AND " &
                        ClsIdDoc_NotaReversaCrEnt.SstrNombreCampoBd & " = " & lentIdRC
                Dim ldtbNotasRCr = ClsPanorama.FdtbDataTable(
                        ClsNotaReversionCr.SstrNombreTabla,
                        lstrCamposSelect, {{}}, lstrFiltro)
                If ldtbNotasRCr.Rows.Count > 0 Then
                    Dim lstrPref As String = ClsPanorama.FobjValorCampo(ldtbNotasRCr(0)(0),
                            EnuTipoValor.EnuString)
                    Dim lentId As Integer = ClsPanorama.FobjValorCampo(ldtbNotasRCr(0)(1),
                            EnuTipoValor.EnuInteger)
                    MstrNroNotaRCr = ClsPanorama.FstrNumeroDcto(lstrPref, lentId)
                End If
            End If
            Return MstrNroNotaRCr
        End Get
    End Property
    Friend ReadOnly Property ObjNotaReversionCr As ClsNotaReversionCr
        Get
            If IsNothing(MobjNotaReversionRC) AndAlso StrNroNotaRCr <> "" Then
                Dim lstrPref As String = ClsPanorama.FstrPrefijoDcto(StrNroNotaRCr)
                Dim lentId As Integer = ClsPanorama.FentIdDcto(StrNroNotaRCr)
                Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentId}
                MobjNotaReversionRC = New ClsNotaReversionCr()
                MobjNotaReversionRC.SAbra(lobjValorLlave)
            End If
            Return MobjNotaReversionRC
        End Get
    End Property
#End Region

#Region "Descuentos"
    Friend StrNroFacDscto As String
    Friend ShrIdItemFacDscto As Short
    Friend EnuTipoDscto As EnuTipoDescuento
    Friend StrVlrDscto As String
    Friend ReadOnly Property DtbDescuentos() As DataTable
        Get
            If MdtbDescuentos Is Nothing AndAlso MobjClienteRec IsNot Nothing AndAlso
                    MobjClienteRec.BlnExiste AndAlso ObjServicios_RecStr.BlnEsValido Then
                SCargueDsctos()
            End If
            Return MdtbDescuentos
        End Get
    End Property
    Private Sub SCargueDsctos()
        Dim lstrIdpredsAgru = ObjIdPredioAgrupador_RecStr.ToString().Split(",")
        Dim lstrServicios = ObjServicios_RecStr.ToString().Split(",")
        MdtbDescuentos = MobjClienteRec.FdtbDescuentos(lstrIdpredsAgru, lstrServicios,
                ObjFechaRecDtm.ObjValorPro)
    End Sub
    ''' <summary>
    ''' Devuelve el total de descuentos a efectuarse en el momento actual
    ''' </summary>   
    Friend ReadOnly Property DecTotalDsctos As Decimal
        Get
            Dim ldecTotalDsctos = 0D
            If Not IsNothing(MdtbDescuentos) Then
                Dim ldrwDescuentos As DataRow() = MdtbDescuentos.Select()
                For Each ldrwDscto As DataRow In ldrwDescuentos
                    ldecTotalDsctos += ldrwDscto("Valor")
                Next
            End If
            Return ldecTotalDsctos
        End Get
    End Property
    Friend Function FblnHayDscto() As Boolean
        Dim lblnHay = False
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando AndAlso
                DtbDescuentos IsNot Nothing Then
            lblnHay = DtbDescuentos IsNot Nothing AndAlso DtbDescuentos.Rows.Count > 0
        End If
        Return lblnHay
    End Function
    Private Function FblnHayDscto(astrNroFact As String, ashrIdItemFac As Short,
             aenuTipoDscto As EnuTipoDescuento) As Boolean
        Dim lstrNroFac As String, lenuTipoDscto As EnuTipoDescuento
        Dim lstrItemFac As String, lshrIdItemFac As Short
        Dim lblnHayDscto = False
        For Each ldrwDscto As DataRow In DtbDescuentos.Rows
            lstrNroFac = ClsPanorama.FobjValorCampo(ldrwDscto("NroFactura"), EnuTipoValor.EnuString)
            lstrItemFac = ClsPanorama.FobjValorCampo(ldrwDscto("IdItemFact"), EnuTipoValor.EnuString)
            lshrIdItemFac = lstrItemFac.Substring(0, lstrItemFac.IndexOf("-"))
            lenuTipoDscto = ClsPanorama.FobjValorCampo(ldrwDscto("IdTipoDcto"), EnuTipoValor.EnuInteger)
            lblnHayDscto = (lstrNroFac = astrNroFact) AndAlso (lshrIdItemFac = ashrIdItemFac) AndAlso
                    (lenuTipoDscto = aenuTipoDscto)
            If lblnHayDscto Then Exit For
        Next
        Return lblnHayDscto
    End Function
    Friend Function FblnHayDsctoPP() As Boolean
        Dim lblnHay = False
        If Not IsNothing(MdtbDescuentos) Then
            Dim lenuTipoDsct As EnuTipoDescuento
            For Each ldrwDscto In MdtbDescuentos.Rows
                lenuTipoDsct = ClsPanorama.FobjValorCampo(ldrwDscto("IdTipoDcto"), EnuTipoValor.EnuInteger)
                If lenuTipoDsct = EnuTipoDescuento.EnuDsctoPP Then
                    lblnHay = True
                    Exit For
                End If
            Next
        End If
        Return lblnHay
    End Function
    Friend Sub SAdicionesDscto(astrNroFactura As String, astrItemFac As String,
                 aenuTipoDscto As EnuTipoDescuento, adecvalorDscto As Decimal)
        Dim ldecBaseDscto = 0D
        Dim lshrIdItemFac As Short = astrItemFac.Substring(0, astrItemFac.IndexOf("-"))
        Dim ldblTasaDscto = ClsOrionCop.FdblTasaDscto(astrNroFactura, lshrIdItemFac,
                aenuTipoDscto, adecvalorDscto, ldecBaseDscto)
        If IsNothing(MdtbDescuentos) Then
            MdtbDescuentos = ClsOrionCop.FdtbDescuentos
        End If
        Dim ldrwNuevoDscto As DataRow = MdtbDescuentos.NewRow
        Dim lentOrdinal = MdtbDescuentos.Rows.Count + 1
        ldrwNuevoDscto("Ordinal") = lentOrdinal
        ldrwNuevoDscto("NroFactura") = astrNroFactura
        ldrwNuevoDscto("IdItemFact") = astrItemFac
        ldrwNuevoDscto("IdTipoDcto") = aenuTipoDscto
        ldrwNuevoDscto("TipoDcto") = ClsOrionCop.FstrNombreDatoConstanteOri(EnuGrupoConstantesOriDef.EnuTipoDescuento, aenuTipoDscto)
        ldrwNuevoDscto("Base") = ldecBaseDscto
        ldrwNuevoDscto("Tasa") = ldblTasaDscto
        ldrwNuevoDscto("Valor") = adecvalorDscto
        MdtbDescuentos.Rows.Add(ldrwNuevoDscto)
    End Sub
    Friend Sub SElimineDscto(aentOrdinal As Integer)
        If MdtbDescuentos.Rows.Count > 0 Then
            If aentOrdinal <= MdtbDescuentos.Rows.Count Then
                MdtbDescuentos.Rows(aentOrdinal - 1).Delete()
                If MdtbDescuentos.Rows.Count > 0 Then
                    For i = 0 To MdtbDescuentos.Rows.Count - 1
                        Dim ldrwDscto As DataRow = MdtbDescuentos.Rows(i)
                        ldrwDscto("Ordinal") = i + 1
                    Next
                End If
            End If
        End If
    End Sub
    Friend Function FblnEsValidoDscto() As Boolean
        Dim lstrMens = String.Empty, lblnEsValido As Boolean
        Dim lenuSevNot As EnuSeveridadNot = EnuSeveridadNot.None
        lblnEsValido = FblnEsValidoDescuento(lstrMens, lenuSevNot)
        If Not lblnEsValido OrElse lenuSevNot = EnuSeveridadNot.EnuAdvertencia Then
            SLevanteEventoNot(lstrMens, "", EnuIdMens.EnuDscto, lenuSevNot)
        Else
            SLevanteEventoNot(lstrMens, "", EnuIdMens.EnuDscto, lenuSevNot)
        End If
        Return lblnEsValido
    End Function

    Private Function FblnEsValidoDescuento(ByRef astrMensaje As String,
            ByRef aenuSevNoti As EnuSeveridadNot) As Boolean
        Dim lblnEsValido As Boolean, ldecVlrDscto As Decimal, lstrMens = String.Empty
        Dim lenuSevNot As EnuSeveridadNot = EnuSeveridadNot.None
        If Not String.IsNullOrEmpty(StrVlrDscto) Then
            If IsNumeric(StrVlrDscto) Then
                ldecVlrDscto = CType(StrVlrDscto, Decimal)
                lblnEsValido = (ldecVlrDscto > 0)
            Else
                lblnEsValido = False
                lstrMens = "El Dato debe ser Numérico!"
                lenuSevNot = EnuSeveridadNot.EnuDatoInvalido
            End If
            If lblnEsValido Then
                lblnEsValido = FblnEsValidoDscto(StrNroFacDscto, ShrIdItemFacDscto,
                        EnuTipoDscto, ldecVlrDscto, lstrMens, lenuSevNot)
            End If
        Else
            lblnEsValido = True
        End If
        aenuSevNoti = lenuSevNot
        astrMensaje = lstrMens
        Return lblnEsValido
    End Function
    Private Function FblnEsValidoDscto(astrNroFactura As String, ashrIdItemFac As Short,
             aenuTipoDescuento As EnuTipoDescuento, adecValorDscto As Decimal,
             ByRef astrMensaje As String, ByRef aenuSevNoti As EnuSeveridadNot)
        Dim lstrMens = String.Empty, lenuSevNot As EnuSeveridadNot
        Dim lblnEsValido As Boolean
        lblnEsValido = adecValorDscto > 0
        If Not lblnEsValido Then
            lstrMens = "El valor del descuento debe ser mayor a cero!"
        End If
        If lblnEsValido Then
            lblnEsValido = (aenuTipoDescuento <> EnuTipoDescuento.EnuCancelaIva)
            If Not lblnEsValido Then
                lstrMens = "No es posible llevar el IVA al Gasto desde un Recibo de Caja!"
            Else
                lblnEsValido = adecValorDscto - Int(adecValorDscto) = 0
                If Not lblnEsValido Then
                    lstrMens = "El Valor ingresado debe ser sin Centavos!"
                End If
            End If
        End If
        If lblnEsValido Then
            If GobjParametros.ObjAutorizaEFacBln.ObjValorPro Then
                Dim lstrPref = ClsPanorama.FstrPrefijoDcto(astrNroFactura)
                Dim lentIdFac = ClsPanorama.FentIdDcto(astrNroFactura)
                lblnEsValido = MobjClienteRec.FblnFactEstadoEFacOk(lstrPref, lentIdFac)
                If Not lblnEsValido Then
                    lstrMens = "No es posible un Descuento a una Factura no registrada!"
                End If
            End If
        End If
        If lblnEsValido Then
            ' Solo puede existir un descuento por item de factura y por tipo de descuento
            lblnEsValido = Not FblnHayDscto(astrNroFactura, ashrIdItemFac, aenuTipoDescuento)
            lstrMens = "Un Descuento de este tipo, a este Servicio, ya fue ingresado!"
        End If
        If lblnEsValido Then
            lblnEsValido = MdtbDescuentos.Rows.Count <= 15
            lstrMens = "No está permitido hacer más de 15 descuentos en un Recibo de Caja"
        End If
        If Not lblnEsValido Then
            lenuSevNot = EnuSeveridadNot.EnuDatoInvalido
        Else
            Dim lobjFra As New ClsFactura()
            Dim lstrPref = ClsPanorama.FstrPrefijoDcto(astrNroFactura)
            Dim lentIdFac = ClsPanorama.FentIdDcto(astrNroFactura)
            Dim lobjVlrLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentIdFac}
            lobjFra.SAbra(lobjVlrLlave)
            lblnEsValido = lobjFra.FblnEsValidoDescuento(ashrIdItemFac, adecValorDscto,
                    aenuTipoDescuento, lstrMens, lenuSevNot)
        End If
        If lblnEsValido Then
            lstrMens = String.Empty
            If lenuSevNot <> EnuSeveridadNot.EnuAdvertencia Then
                lenuSevNot = EnuSeveridadNot.EnuOk
            End If
        End If
        astrMensaje = lstrMens
        aenuSevNoti = lenuSevNot
        Return lblnEsValido
    End Function
    Friend Sub SVerfiqueDsctoPP()
        Dim ldecSaldoAPagar = FdecSaldoRC()
        If ldecSaldoAPagar > 0 Then
            ' Si el saldo a pagar > 0 reviso descuentos y si hay descuento por pronto pago advierto
            If FblnHayDsctoPP() Then
                Dim lstrMens = "Está aplicando Descuento por Pronto Pago teniendo Deuda!"
                SLevanteEventoNot(lstrMens, "", 0, EnuSeveridadNot.EnuAdvertencia)
            End If
        End If
    End Sub
#End Region

#Region "Reversa Notas Cr"
    Friend Sub SReverseNotasCr(adtmFechaReversion As Date)
        Dim lstrPrefNCr = String.Empty, lentIdNCr = 0
        Dim lobjNotaCr As New ClsNotaCr()
        Dim lobjValorLlave As Object() = Nothing
        Dim lstrPref = GobjParametros.FstrPrefijoDoc(EnuIdDocumentoDef.EnuNotaReversaCr)
        Dim lobjNotaRcr As New ClsNotaReversionCr(lstrPref)
        For Each lstrIdNCr As String In ObjIdNotasCrStr.ToString().Split(",")
            If Not String.IsNullOrEmpty(lstrIdNCr) Then
                lstrPrefNCr = ClsPanorama.FstrPrefijoDcto(lstrIdNCr)
                lentIdNCr = ClsPanorama.FentIdDcto(lstrIdNCr)
                lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrPrefNCr, lentIdNCr}
                lobjNotaCr.SAbra(lobjValorLlave)
                If ObjNotaReversionCr Is Nothing Then
                    If Not lobjNotaCr.ObjAnuladoBln.ObjValorPro Then
                        If Not lobjNotaCr.ObjAnuladoBln.ObjValorPro Then
                            lobjNotaRcr.SCreeObj(Nothing)
                            With lobjNotaRcr
                                .ObjDetalle_NotaReversaCrStr.ObjValorPro = "Anulada por anulacion Recibo Caja Nro. " &
                                    StrNumeroRecCaja
                                .ObjFecha_NotaReversaCrDtm.ObjValorPro = adtmFechaReversion
                                .ObjIdCliente_NotaReversaCrDbl.ObjValorPro = lobjNotaCr.ObjIdCliente_NotaCrDbl.ObjValorPro
                                .ObjIdDoc_NotaReversaCrEnt.ObjValorPro = lentIdNCr
                                .ObjIdPredioAgrupador_NotaReversaCrStr.ObjValorPro =
                                    lobjNotaCr.ObjIdPredioAgrupador_NotaCrStr.ObjValorPro
                                .ObjPrefijoDoc_NotaReversaCrStr.ObjValorPro = lstrPrefNCr
                                .ObjTipoDocReversadoByt.ObjValorPro = EnuDocReversado.EnuNotaCr
                                .ObjValor_NotaReversaCrDec.ObjValorPro = lobjNotaCr.ObjValor_NotaCrDec.ObjValorPro
                                .SActualice(True)
                            End With
                        End If
                    End If
                End If
            End If
        Next
    End Sub
#End Region

#Region "Novedades"
    Friend ReadOnly Property ColNovedades As Collection
        Get
            If IsNothing(McolNovedades) Then
                McolNovedades = New Collection
                SCargueDtbNovedades()
                If Not IsNothing(MdtbNovedades) AndAlso MdtbNovedades.Rows.Count > 0 Then
                    Dim ldrwNovedades() As DataRow = MdtbNovedades.Select
                    For Each ldrwNovedad As DataRow In ldrwNovedades
                        Dim lobjNovedad As New ClsNovedad(Me, ldrwNovedad)
                        lobjNovedad.SLeaValores(True)
                        McolNovedades.Add(lobjNovedad)
                    Next
                End If
            End If
            Return McolNovedades
        End Get
    End Property
    ''' <summary>
    ''' Devuelve un DataTable con las novedades y las novedades del anticipo para ser
    ''' mostradas en la interfaz del Recibo de Caja
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend ReadOnly Property DtbNovedadesRC As DataTable
        Get
            SCargueDtbNovedades()
            Return MdtbNovedades
        End Get
    End Property
    Private Sub SCargueDtbNovedades()
        If IsNothing(MdtbNovedades) Then
            Dim lstrIdRecCaja = ObjIdRecCajaEnt.ToString
            If String.IsNullOrEmpty(lstrIdRecCaja) Then lstrIdRecCaja = "0"
            Dim lstrCamposSelect() = {"*"}
            Dim lstrIndice = {{ClsPrefijoFact_NovStr.SstrNombreCampoBd, "ASC"},
                              {ClsIdFactura_NovEnt.SstrNombreCampoBd, "ASC"},
                              {ClsIdNovedadShr.SstrNombreCampoBd, "ASC"}}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsIdTipoDocOrigenByt.SstrNombreCampoBd &
                    " = " & EnuTipoDocOri.EnuReciboCaja & " AND " &
                    ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd & " = '" & ObjPrefijo_RecStr.ObjValorPro & "' AND " &
                    ClsIdDocOrigenEnt.SstrNombreCampoBd & " = " & lstrIdRecCaja & " AND " &
                    ClsValor_NovDec.SstrNombreCampoBd & " <> 0"
            MdtbNovedades = ClsPanorama.FdtbDataTable(ClsNovedad.SstrNombreTabla, lstrCamposSelect, lstrIndice,
                    lstrFiltro)
        End If
    End Sub
    Friend Function FdtbNovedadesRC()
        Dim lstrExpSqlNov = FstrExpSqlNovedades()
        Dim lstrExpSqlNovAnt = FstrExpSqlNovedadesAnt()
        Dim lstrIndice = " ORDER BY " & ClsPrefijoFact_NovStr.SstrNombreCampoBd & ", " &
                ClsIdFactura_NovEnt.SstrNombreCampoBd & ", " & ClsIdNovedadShr.SstrNombreCampoBd
        Dim lstrSql = (lstrExpSqlNov & " UNION ALL " & lstrExpSqlNovAnt) + lstrIndice
        Dim ldtbNovedadesRC = ClsPanorama.FdtbDataTable(lstrSql)
        SComplementeTablaNov(ldtbNovedadesRC)
        Return ldtbNovedadesRC
    End Function
    Private Function FstrExpSqlNovedades() As String
        Dim lstrIdRecCaja = ObjIdRecCajaEnt.ToString
        If String.IsNullOrEmpty(lstrIdRecCaja) Then lstrIdRecCaja = "0"
        Dim lstrTabla = ClsNovedad.SstrNombreTabla
        Dim lstrCamposSelect = {ClsPrefijoFact_NovStr.SstrNombreCampoBd,
                                ClsIdFactura_NovEnt.SstrNombreCampoBd,
                                "'' as NroFac",
                                ClsIdItemFacturaShr.SstrNombreCampoBd,
                                ClsIdNovedadShr.SstrNombreCampoBd,
                                "'' AS Detalle",
                                ClsIdCuentaDb_NovStr.SstrNombreCampoBd,
                                ClsIdCuentaCr_NovStr.SstrNombreCampoBd,
                                ClsValor_NovDec.SstrNombreCampoBd,
                                ClsIdTipoNovedadByt.SstrNombreCampoBd}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                        ClsIdTipoDocOrigenByt.SstrNombreCampoBd & " = " &
                        EnuTipoDocOri.EnuReciboCaja & " AND " &
                        ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd & " = '" &
                        ObjPrefijo_RecStr.ObjValorPro & "' AND " &
                        ClsIdDocOrigenEnt.SstrNombreCampoBd & " = " &
                        lstrIdRecCaja & " AND " & ClsValor_NovDec.SstrNombreCampoBd & " <> 0"
        Dim lstrIndice = {{"", ""}}
        Dim lstrSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTabla, lstrCamposSelect,
                lstrIndice, lstrFiltro, Array.Empty(Of String))
        Return lstrSql
    End Function
    Private Function FstrExpSqlNovedadesAnt() As String
        Dim lstrIdRecCaja = ObjIdRecCajaEnt.ToString
        If String.IsNullOrEmpty(lstrIdRecCaja) Then lstrIdRecCaja = "0"
        Dim lstrTabla = ClsNovedadAnticipo.SstrNombreTabla
        Dim lstrCamposSelect = {"'' AS " & ClsPrefijoFact_NovStr.SstrNombreCampoBd,
                                "0 AS " & ClsIdFactura_NovEnt.SstrNombreCampoBd,
                                "'' as NroFac",
                                "0 AS " & ClsIdItemFacturaShr.SstrNombreCampoBd,
                                ClsIdNovedadAntShr.SstrNombreCampoBd,
                                "'' AS Detalle",
                                ClsIdCuentaDb_NovAntStr.SstrNombreCampoBd,
                                ClsIdCuentaCr_NovAntStr.SstrNombreCampoBd,
                                ClsValor_NovAntDec.SstrNombreCampoBd,
                                ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                        ClsIdTipoDocOrigen_AntByt.SstrNombreCampoBd & " = " &
                        EnuTipoDocOri.EnuReciboCaja & " AND " &
                        ClsPrefijoDocOrigen_NovAntStr.SstrNombreCampoBd & " = '" &
                        ObjPrefijo_RecStr.ObjValorPro & "' AND " &
                        ClsIdDocOrigen_NovAntEnt.SstrNombreCampoBd & " = " &
                        lstrIdRecCaja & " AND " & ClsValor_NovAntDec.SstrNombreCampoBd & " <> 0"
        Dim lstrIndice = {{"", ""}}
        Dim lstrSql = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTabla, lstrCamposSelect,
                lstrIndice, lstrFiltro, Array.Empty(Of String))
        Return lstrSql
    End Function
    Private Shared Sub SComplementeTablaNov(adtbNovedadesRC As DataTable)
        Dim ldrwNovedades = adtbNovedadesRC.Select
        Dim lstrConceptoNovedad = String.Empty, lstrPrefFac = String.Empty
        Dim lentIdFac = 0, lstrNroFac = String.Empty
        Dim lenuTipoNovedad As EnuTipoNov = EnuTipoNov.None
        For Each ldrwNovedad As DataRow In ldrwNovedades
            lstrPrefFac = ClsPanorama.FobjValorCampo(ldrwNovedad(ClsPrefijoFact_NovStr.SstrNombreCampoBd),
                    EnuTipoValor.EnuString)
            lentIdFac = ClsPanorama.FobjValorCampo(ldrwNovedad(ClsIdFactura_NovEnt.SstrNombreCampoBd),
                    EnuTipoValor.EnuInteger)
            lstrNroFac = ClsPanorama.FstrNumeroDcto(lstrPrefFac, lentIdFac)
            lenuTipoNovedad = ClsPanorama.FobjValorCampo(ldrwNovedad(ClsIdTipoNovedadByt.SstrNombreCampoBd),
                    EnuTipoValor.EnuByte)
            Select Case lenuTipoNovedad
                Case EnuTipoNov.EnuCrPagoCap
                    lstrConceptoNovedad = My.Resources.RCAbonoCap
                Case EnuTipoNov.EnuCrPagoInt
                    lstrConceptoNovedad = My.Resources.RCAbonoMor
                Case EnuTipoNov.EnuCrDctoCap
                    lstrConceptoNovedad = My.Resources.RCDsctoCap
                Case EnuTipoNov.EnuCrDctoInt
                    lstrConceptoNovedad = My.Resources.RCDsctoMor
                Case EnuTipoNov.EnuCrRetFte
                    lstrConceptoNovedad = My.Resources.RCReteFte
                Case EnuTipoNov.EnuCrRetIca
                    lstrConceptoNovedad = My.Resources.RCRetIca
                Case EnuTipoNov.EnuCrRetIva
                    lstrConceptoNovedad = My.Resources.RCRetIva
                Case EnuTipoNov.EnuCrAntRec
                    lstrConceptoNovedad = My.Resources.RCAntRec
                Case EnuTipoNov.EnuRCrPagoCap
                    lstrConceptoNovedad = My.Resources.RCRAbonoCap
                Case EnuTipoNov.EnuRCrPagoInt
                    lstrConceptoNovedad = My.Resources.RCRAbonoMor
                Case EnuTipoNov.EnuRCrDctoCap
                    lstrConceptoNovedad = My.Resources.RCRDsctoCap
                Case EnuTipoNov.EnuRCrDctoInt
                    lstrConceptoNovedad = My.Resources.RCRDsctoMor
                Case EnuTipoNov.EnuRCrRetFte
                    lstrConceptoNovedad = My.Resources.RCRReteFte
                Case EnuTipoNov.EnuRCrRetIca
                    lstrConceptoNovedad = My.Resources.RCRRetIca
                Case EnuTipoNov.EnuRCrRetIva
                    lstrConceptoNovedad = My.Resources.RCRRetIva
                Case EnuTipoNov.EnuRCrAntRec
                    lstrConceptoNovedad = My.Resources.RCRAntRec
            End Select
            ldrwNovedad("NroFac") = lstrNroFac
            ldrwNovedad("Detalle") = lstrConceptoNovedad
        Next
    End Sub
#End Region

#Region "Integridad"
    Private Function FblnEsIntegroRC() As Boolean
        Dim lblnRcOk = True
        Dim ldecVlrAnticipo = 0D, ldecSumaValorItems = 0D
        For Each lobjItemRec As ClsItemRecCaja In ColItemsRecCaja
            Dim ldecValorItem As Decimal = lobjItemRec.ObjValor_ItemRecDec.ObjValorPro
            Dim lenuTipoItemRec As EnuTipoItemRecCajaDef = lobjItemRec.ObjIdTipoItemRecByt.ObjValorPro
            If lenuTipoItemRec <> EnuTipoItemRecCajaDef.EnuAnticipo Then
                lblnRcOk = FblnEsIntegroItemRC(lobjItemRec)
            End If
            If lblnRcOk Then
                Select Case lenuTipoItemRec
                    Case EnuTipoItemRecCajaDef.EnuAbonoCapital, EnuTipoItemRecCajaDef.EnuAbonoIntMora
                        ldecSumaValorItems += ldecValorItem
                    Case EnuTipoItemRecCajaDef.EnuAnticipo
                        ldecSumaValorItems += ldecValorItem
                        ldecVlrAnticipo += ldecValorItem
                End Select
            End If
        Next
        If lblnRcOk Then
            lblnRcOk = (ObjValorAnticipoDec.ObjValorPro = ldecVlrAnticipo)
        End If
        If lblnRcOk Then
            lblnRcOk = (ObjValor_RecDec.ObjValorPro = ldecSumaValorItems)
        End If
        Return lblnRcOk
    End Function
    Private Function FblnEsIntegroItemRC(aobjItemRc As ClsItemRecCaja) As Boolean
        Dim lblnEsIntegro = True
        Dim lenuTipoNov As EnuTipoNov
        Dim lenuTipoItemRec As EnuTipoItemRecCajaDef = aobjItemRc.ObjIdTipoItemRecByt.ObjValorPro
        Dim lblnHayError As Boolean, ldecVlrNovedades = 0D
        For Each lobjNov As ClsNovedad In ColNovedades
            If aobjItemRc.ObjIdItemRecCajaShr.ObjValorPro = lobjNov.ObjIdItemDocOrigen_NovShr.ObjValorPro Then
                lenuTipoNov = lobjNov.ObjIdTipoNovedadByt.ObjValorPro
                ldecVlrNovedades += lobjNov.ObjValor_NovDec.ObjValorPro
                Select Case lenuTipoNov
                    Case EnuTipoNov.EnuCrDctoCap
                        lblnHayError = (lenuTipoItemRec <> EnuTipoItemRecCajaDef.EnuDsctoCapital) AndAlso
                                  (lenuTipoItemRec <> EnuTipoItemRecCajaDef.EnuDsctoPP)
                    Case EnuTipoNov.EnuCrDctoInt
                        lblnHayError = (lenuTipoItemRec <> EnuTipoItemRecCajaDef.EnuDsctoIntMora)
                    Case EnuTipoNov.EnuCrPagoCap
                        lblnHayError = (lenuTipoItemRec <> EnuTipoItemRecCajaDef.EnuAbonoCapital)
                    Case EnuTipoNov.EnuCrPagoInt
                        lblnHayError = (lenuTipoItemRec <> EnuTipoItemRecCajaDef.EnuAbonoIntMora)
                    Case EnuTipoNov.EnuCrRetFte
                        lblnHayError = (lenuTipoItemRec <> EnuTipoItemRecCajaDef.EnuReteFuente)
                    Case EnuTipoNov.EnuCrRetIca
                        lblnHayError = (lenuTipoItemRec <> EnuTipoItemRecCajaDef.EnuReteIca)
                    Case EnuTipoNov.EnuCrRetIva
                        lblnHayError = (lenuTipoItemRec <> EnuTipoItemRecCajaDef.EnuReteIva)
                    Case Else
                        lblnHayError = True
                End Select
                If lblnHayError Then
                    lblnEsIntegro = False
                    Exit For
                End If
            End If
        Next
        If lblnEsIntegro Then
            lblnEsIntegro = (ldecVlrNovedades = aobjItemRc.ObjValor_ItemRecDec.ObjValorPro)
        End If
        Return lblnEsIntegro
    End Function
#End Region

#Region "Notificaciones"
    Friend Overrides Function FblnNotificaOk(aenuIdMensNot As EnuIdMens) As Boolean
        Static lblnNotiOk = True
        lblnNotiOk = MyBase.FblnNotificaOk(aenuIdMensNot)
        If lblnNotiOk Then
            Select Case aenuIdMensNot
                Case EnuIdMens.EnuNoCreable
                    Dim lstrMens = ""
                    lblnNotiOk = ClsOrionCop.FblnPuedeCrear(EnuTipoDocOri.EnuReciboCaja, False,
                            lstrMens)
                Case EnuIdMens.EnuDscto
                    Dim lstrMens = String.Empty
                    Dim lenuSevNot As EnuSeveridadNot = EnuSeveridadNot.None
                    lblnNotiOk = FblnEsValidoDescuento(lstrMens, lenuSevNot) AndAlso
                            lenuSevNot <> EnuSeveridadNot.EnuAdvertencia
                Case EnuIdMens.EnuTotalMP
                    lblnNotiOk = FdecValorTotalMedPago() <= ObjValor_RecDec.ObjValorPro
            End Select
        End If
        Return lblnNotiOk
    End Function
#End Region

#Region "eFac"
    ''' <summary>
    ''' Indica si el descuento en el recibo de caja  esta haciendo una devolución total del valor
    ''' de la factura afectada.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>Aplica para descuentos en recibo de caja cuando esta instalado el módulo de 
    ''' facturación electrónica</remarks>
    Private Function FblnEsDesctoTotal() As Boolean
        Dim lblnEsDsctoTotal = False, ldecTotalDscto = 0D
        Dim lenuTipoItResCaj As EnuTipoItemRecCajaDef
        Dim lstrPrefFac = String.Empty, lentIdFac = 0
        For Each lobjItemRC As ClsItemRecCaja In ColItemsRecCaja
            lenuTipoItResCaj = lobjItemRC.ObjIdTipoItemRecByt.ObjValorPro
            If lenuTipoItResCaj = EnuTipoItemRecCajaDef.EnuDsctoCapital OrElse
                    lenuTipoItResCaj = EnuTipoItemRecCajaDef.EnuDsctoIntMora OrElse
                    lenuTipoItResCaj = EnuTipoItemRecCajaDef.EnuDsctoPP Then
                ldecTotalDscto += lobjItemRC.ObjValor_ItemRecDec.ObjValorPro
                If lstrPrefFac = String.Empty AndAlso lentIdFac = 0 Then
                    lstrPrefFac = lobjItemRC.ObjPrefijoFact_ItemRecStr.ObjValorPro
                    lentIdFac = lobjItemRC.ObjIdFactura_ItemRecEnt.ObjValorPro
                End If
            End If
        Next
        If ldecTotalDscto > 0 Then
            Dim lobjFra As New ClsFactura()
            Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, lstrPrefFac, lentIdFac}
            lobjFra.SAbra(lobjValorLlave)
            If lobjFra.ObjDebitos_FactDec.ObjValorPro = lobjFra.ObjCreditos_FactDec.ObjValorPro Then
                lblnEsDsctoTotal = (lobjFra.ObjDebitos_FactDec.ObjValorPro = ldecTotalDscto)
            End If
        End If
        Return lblnEsDsctoTotal
    End Function
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsComentario_RecStr
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsReciboCaja = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "Comentario"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Comentario"
        HshrLongitud = 500
        HenuTipoValor = EnuTipoValor.EnuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 3, ShrLongitud,
                    BlnEsRequerido)
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
            HblnEsRequerido = MobjPadre.FblnHayDscto
            HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 3, ShrLongitud,
                    BlnEsRequerido)
            If Not HblnEsValido Then
                HstrMens = "Es indispensable un Comentario cuando hay Descuentos!"
                SNotifiqueDatInv()
            End If
            If Not String.IsNullOrEmpty(HobjValorNew.ToString) Then
                HobjValorNew = HobjValorNew.ToString.Replace(",", ".")
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
            Return HobjValorPro
        Else
            Return ""
        End If
    End Function
End Class

Friend Class ClsRCEnviadoMailBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "EnviadoMail"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Enviada por Email"
        HenuTipoValor = EnuTipoValor.EnuBoolean
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

Friend Class ClsFechaAnulacion_RecDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaAnulacion"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "FechaAnulacion_Rec"
        HenuTipoValor = EnuTipoValor.EnuDateTime
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = False
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MyBase.SVaciePropiedad()
        HobjValorNew = GCDTMFECHANULA
        HobjValorPro = HobjValorNew
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsReciboCaja = ObjPadre
        HblnEsValido = Not IsNothing(HobjValorNew) AndAlso IsDate(HobjValorNew)
        If HobjValorNew <> GCDTMFECHANULA Then
            If HblnEsValido Then
                Dim ldtmFechaMin As Date = GCDTMFECHANULA
                Dim ldtmFechaMax As Date = GCDTMFECHANULA
                If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando Then
                    HblnEsValido = lobjPadre.FblnEsAnulable
                    If HblnEsValido Then
                        HblnEsValido = lobjPadre.ObjAnuladoBln.ObjValorPro
                        If HblnEsValido Then
                            If Today > GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo Then
                                ldtmFechaMin = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
                                ldtmFechaMax = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
                            Else
                                ldtmFechaMin = Date.Today
                                ldtmFechaMax = Now
                            End If
                        End If
                        HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax,
                                BlnEsRequerido)
                    End If
                ElseIf lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
                    HblnEsValido = HobjValorNew = HobjValorOriginal
                Else
                    HblnEsValido = False
                End If
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Private Sub EPosSetValor() Handles Me.EvnPosSetValor
        Dim lobjPadre As ClsReciboCaja = ObjPadre
        If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.EnuConsultando AndAlso BlnEsValido Then
            If HobjValorPro <> GCDTMFECHANULA Then
                lobjPadre.ObjValor_RecDec.SValide()
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

Friend Class ClsFechaRecDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaReciboCaja"
    Private ReadOnly MobjPadre As ClsReciboCaja = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "FechaReciboCaja"
        HenuTipoValor = EnuTipoValor.EnuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MyBase.SVaciePropiedad()
        If GobjParametros.ObjAnoActual.StrIdPeriodoActual = ClsOrionCop.FstrPeriodoDeFecha(Date.Today) Then
            HobjValorNew = Date.Today
            HobjValorPro = GCDTMFECHANULA
        Else
            HobjValorNew = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
            HobjValorPro = GCDTMFECHANULA
        End If
        HobjValorOriginal = GCDTMFECHANULA
    End Sub
    Public Overrides Sub SValide()
        Dim ldtmFechaMin = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
        Dim ldtmFechaMax = Date.Today()
        Dim lstrPeriodoHoy = ClsPanorama.FstrPeriodo(Date.Today)
        Dim lstrPeriodoActual = GobjParametros.ObjAnoActual.StrIdPeriodoActual
        If lstrPeriodoActual < lstrPeriodoHoy Then
            ldtmFechaMax = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
        End If
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
            HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax, BlnEsRequerido)
            If Not HblnEsValido AndAlso IsDate(HobjValorNew) Then
                If HobjValorNew > Date.Today Then
                    HstrMens = "La Fecha es posterior al Día de Hoy!"
                Else
                    HstrMens = "La Fecha está por fuera del Período Actual!"
                End If
                SNotifiqueDatInv()
            Else
                If MobjPadre.ObjIdPredioAgrupador_RecStr.BlnEsValido Then
                    SValideFechaUltimaNov()
                End If
            End If
        Else
            HblnEsValido = HobjValorNew = HobjValorOriginal
        End If
    End Sub
    Private Sub EPosCambioVlr(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosCambio
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
            MobjPadre.SVerifiqueDeuda()
        End If
    End Sub
    ''' <summary>
    ''' Valida que la fecha del recibo no sea anterior a la ultima novedad de 
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub SValideFechaUltimaNov()
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
            Dim ldtmFechaUltimaNov = GCDTMFECHANULA
            Dim lblnEsValido = False
            Dim lobjCliente As ClsCliente = MobjPadre.ObjClienteRecibo
            If Not IsNothing(lobjCliente) AndAlso lobjCliente.BlnExiste Then
                Dim lstrPresAgr As String = MobjPadre.ObjIdPredioAgrupador_RecStr.ObjValorPro
                For Each lstrPreA As String In lstrPresAgr.Split(",")
                    ldtmFechaUltimaNov = lobjCliente.FdtmFechaUltimaNovedad(lstrPreA)
                    lblnEsValido = (HobjValorNew >= ldtmFechaUltimaNov)
                    If Not lblnEsValido Then Exit For
                Next
                If Not lblnEsValido Then
                    Dim lstrUltimaFechaNov = Format(ldtmFechaUltimaNov, GCSTRFMTFECHASIMPLE)
                    HstrMens = "La Fecha del Recibo de Caja no puede ser anterior a la Fecha de la última " &
                            "Novedad del Cliente " & lstrUltimaFechaNov & "!"
                    SNotifiqueDatInv()
                End If
            End If
            HblnEsValido = HblnEsValido AndAlso lblnEsValido
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

Friend Class ClsIdAnticipo_RecEnt
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsReciboCaja = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "IdAnticipo"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Id Anticipo"
        HenuTipoValor = EnuTipoValor.EnuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
            If IsNothing(HobjValorNew) Then HobjValorNew = 0
            HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, Integer.MaxValue,
                    BlnEsRequerido, HenuTipoValor)
        Else
            HblnEsValido = HobjValorNew = HobjValorOriginal
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

Friend Class ClsIdCliente_RecDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTerceroCliente"
    Private ReadOnly MobjPadre As ClsReciboCaja = Nothing
    Private MobjCliente As ClsCliente = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdTerceroCliente_ReciboCaja"
        HenuTipoValor = EnuTipoValor.EnuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = Not IsNothing(HobjValorNew)
        HstrMens = String.Empty
        If HblnEsValido Then
            HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCDBLMINTERC, GCDBLMAXTERC,
                    BlnEsRequerido)
            If HblnEsValido Then
                If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.EnuConsultando Then
                    If IsNothing(MobjCliente) OrElse MobjCliente.ObjIdClienteDbl.ObjValorPro <>
                            HobjValorNew Then
                        MobjCliente = New ClsCliente(EnuModoInstanciaObjDef.EnuUnico)
                    End If
                    If HobjValorNew <> HobjValorPro OrElse Not MobjCliente.BlnExiste Then
                        Dim lobjLlavePrincipal() = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
                        MobjCliente.SAbra(lobjLlavePrincipal)
                    End If
                    If Not MobjCliente.BlnExiste Then
                        HblnEsValido = False
                        MobjCliente.SVacie()
                        MobjPadre.ObjClienteRecibo = Nothing
                        HstrMens = "La Id. del Cliente no es valida!"
                        SNotifiqueDatInv()
                    Else
                        MobjPadre.ObjClienteRecibo = MobjCliente
                        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando AndAlso
                                MobjPadre.ObjIdPredioAgrupador_RecStr.BlnEsValido Then
                            If MobjCliente.FblnEstadoEFacOk(
                                    MobjPadre.ObjIdPredioAgrupador_RecStr.ObjValorPro) Then
                                Dim lenuEstadoDeuda As EnuEstadoDeudaDef =
                                    MobjCliente.FenuEstadoDeuda(MobjPadre.ObjIdPredioAgrupador_RecStr.ObjValorPro)
                                If lenuEstadoDeuda = EnuEstadoDeudaDef.EnuPrejuridico Then
                                    HstrMens = "La Deuda está en estado Pre-Jurídico!"
                                    SLevanteEveNot("", 0, EnuSeveridadNot.EnuAdvertencia)
                                ElseIf lenuEstadoDeuda = EnuEstadoDeudaDef.EnuJuridico Then
                                    HstrMens = "La Deuda está en estado Jurídico!"
                                    SLevanteEveNot("", 0, EnuSeveridadNot.EnuAdvertencia)
                                End If
                            Else
                                HblnEsValido = False
                                HstrMens = "El Cliente tiene facturas electrónicas no registradas!"
                                SLevanteEveNot("", 0, EnuSeveridadNot.EnuInformacion)
                            End If
                        End If
                    End If
                End If
            Else
                MobjCliente = Nothing
                MobjPadre.ObjClienteRecibo = MobjCliente
                HstrMens = "La Id. del Cliente no es valida!"
                SNotifiqueDatInv()
            End If
        End If
    End Sub
    Private Sub EPosCambioVlr(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosCambio
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
            MobjPadre.SVerifiqueDeuda()
        End If
    End Sub
    Friend ReadOnly Property StrNombreCliente As String
        Get
            If Not IsNothing(MobjCliente) AndAlso MobjCliente.BlnExiste Then
                Return MobjCliente.ObjNombreCompletoStr.ObjValorPro
            Else
                Return ""
            End If
        End Get
    End Property
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return String.Empty
        Else
            If HobjValorPro = 0 Then
                Return String.Empty
            Else
                Return HobjValorPro.ToString
            End If
        End If
    End Function
End Class

Friend Class ClsIdNotasCrStr
    'Herencia
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdNotasCr"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Id. Notas Cr"
        HenuTipoValor = EnuTipoValor.EnuString
        HshrLongitud = 80
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud, BlnEsRequerido)
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
            Return String.Empty
        End If
    End Function
End Class

Friend Class ClsIdPredioAgrupador_RecStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdPredioAgrupador"
    Private ReadOnly MobjPadre As ClsReciboCaja = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Predio Agrupador"
        HenuTipoValor = EnuTipoValor.EnuString
        HshrLongitud = 200
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MyBase.SVaciePropiedad()
        HobjValorNew = "***"
        HobjValorPro = "***"
        HobjValorOriginal = "***"
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud,
                BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
                HblnEsValido = HobjValorNew <> "***"
                If HblnEsValido Then
                    If MobjPadre.ObjIdCliente_RecDbl.BlnEsValido Then
                        If MobjPadre.BlnEsSoloAnticipo Then
                            HblnEsValido = HobjValorNew.ToString.Split(",").Length = 1
                            If Not HblnEsValido Then
                                HstrMens = "No es posible un Anticipo a dos Predios Agrupadores!"
                                SNotifiqueDatInv()
                            Else
                                Dim lstrIdPredsAgru As String() = HobjValorNew.ToString().Split(",")
                                HblnEsValido =
                                        MobjPadre.ObjClienteRecibo.FdecDeudaMora(lstrIdPredsAgru) = 0
                                If Not HblnEsValido Then
                                    HstrMens = "No es posible un Anticipo cuando hay " &
                                            "Intereses por pagar!"
                                    SNotifiqueDatInv()
                                End If
                                If MobjPadre.ObjServicios_RecStr.ToString().Contains("0") Then
                                    If HobjValorNew.ToString.Split(",").Contains(String.Empty) Then
                                        HblnEsValido = False
                                        HstrMens = "No se puede hacer un Anticipo por Administración " &
                                                " sin Predio Agrupador"
                                        SNotifiqueDatInv()
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If Not GblnActualizandoApp Then
                    HblnEsValido = HobjValorNew = HobjValorOriginal
                End If
            End If
        Else
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
                If TypeOf HobjValorNew Is String AndAlso HobjValorNew <> "***" AndAlso
                            HobjValorNew.ToString.Length > ShrLongitud Then
                    HstrMens = "La longitud en caracteres de los Predios Agrupadores" &
                                    " es mayor a lo permitido!"
                    SNotifiqueDatInv()
                End If
            End If
        End If
    End Sub
    Private Sub EPosCambioVlr(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosCambio
        Dim lobjPadre As ClsReciboCaja = ObjPadre
        If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
            If HblnEsValido Then
                If lobjPadre.ObjIdCliente_RecDbl.BlnEsValido Then
                    If Not GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro Then
                        lobjPadre.ObjFechaRecDtm.SValideFechaUltimaNov()
                    End If
                End If
                lobjPadre.ObjIdCliente_RecDbl.SValide()
            End If
            If MobjPadre.ObjServicios_RecStr.BlnEsValido Then
                MobjPadre.SVerifiqueDeuda()
            End If
        End If
    End Sub
    Friend Overrides Function FblnNotiInfoOk(aenuIdMensNot As PanL.EnuIdMens) As Boolean
        Dim lblnNotInfOk = True
        If aenuIdMensNot = EnuIdMens.EnuDeudaNoPagable Then
            If MobjPadre.ObjIdCliente_RecDbl.BlnEsValido Then
                If HobjValorNew IsNot Nothing Then
                    For Each lstrIdpreAgr As String In HobjValorNew.Split(",")
                        lblnNotInfOk = MobjPadre.ObjClienteRecibo.FblnPuedeCausarMora(lstrIdpreAgr)
                        If Not lblnNotInfOk Then
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
        Return lblnNotInfOk
    End Function
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

Friend Class ClsIdRecCajaEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdReciboCaja"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdReciboCaja"
        HenuTipoValor = EnuTipoValor.EnuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 3
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue, BlnEsRequerido,
                EnuTipoValor)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return String.Empty
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsPrefijo_RecStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Prefijo"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Prefijo ReciboCaja"
        HshrLongitud = 5
        HenuTipoValor = EnuTipoValor.EnuString
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
            Dim lobjPadre As ClsCBObjetoPan = ObjPadre
            If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
                HblnEsValido = (HobjValorNew =
                        GobjParametros.FstrPrefijoDoc(EnuTipoDocOri.EnuReciboCaja))
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
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsSaldo_RecDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Saldo"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Saldo Cuenta Rec"
        HenuTipoValor = EnuTipoValor.EnuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        Dim lobjPadre As ClsReciboCaja = ObjPadre
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0,
                Decimal.MaxValue, BlnEsRequerido, HenuTipoValor)
        If Not HblnEsValido Then
            If HobjValorNew = 0 Then
                HblnEsValido = lobjPadre.ObjAnuladoBln.ObjValorPro
            End If
        Else
            If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.EnuCreando Then
                If Not GblnActualizandoApp Then
                    HblnEsValido = (HobjValorNew <= lobjPadre.ObjValorDeudaAlPagoDec.ObjValorPro)
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

Friend Class ClsServicios_RecStr
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsReciboCaja = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "ServiciosAfectados"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Servicios Afectados"
        HshrLongitud = 50
        HblnEsRequerido = True
        HenuTipoValor = EnuTipoValor.EnuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            If HobjValorNew.ToString().Contains("A") Then
                If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
                    If HblnEsValido Then
                        If MobjPadre.BlnEsSoloAnticipo Then
                            HblnEsValido = HobjValorNew.ToString().Length = 1
                            If Not HblnEsValido Then
                                HstrMens = "Si está seleccionado <Todos>, no se pueden seleccionar " &
                                    "otros servicio!"
                                SNotifiqueDatInv()
                            End If
                        Else
                            If HobjValorNew.ToString().Contains("A") Then
                                HobjValorNew = "A"
                            End If
                        End If
                    End If
                Else
                    HblnEsValido = MobjPadre.ObjValor_RecDec.ObjValorPro =
                            MobjPadre.ObjValorAnticipoDec.ObjValorPro
                End If
            End If
        End If
    End Sub
    Private Sub EPosCambioVlr(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosCambio
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
            MobjPadre.SVerifiqueDeuda()
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(HobjValorPro) Then
            Dim larlSer As New ArrayList, lstrServicios = String.Empty, i = 0, lblnUltimo As Boolean
            For Each lstrSer In HobjValorPro.ToString().Split(",")
                larlSer.Add(lstrSer)
            Next
            larlSer.Sort()
            For Each lstrSer As String In larlSer
                i += 1
                lstrServicios &= lstrSer
                lblnUltimo = i = larlSer.Count
                If Not lblnUltimo Then
                    lstrServicios &= ","
                End If
            Next
            Return lstrServicios
        Else
            Return ""
        End If
    End Function
End Class

Friend Class ClsValor_RecDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Valor"
    Private ReadOnly MobjPadre As ClsReciboCaja = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Valor Rec"
        HenuTipoValor = EnuTipoValor.EnuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0.01, Decimal.MaxValue,
                BlnEsRequerido, HenuTipoValor)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.EnuCreando Then
                HblnEsValido = HobjValorNew = HobjValorOriginal
            Else
                HblnEsValido = HobjValorNew - Int(HobjValorNew) = 0
                If Not HblnEsValido Then
                    HstrMens = "El Valor ingresado debe ser sin Centavos!"
                    SNotifiqueDatInv()
                Else
                    SValidePago()
                End If
            End If
        Else
            If HobjValorNew = 0 Then
                HblnEsValido = MobjPadre.ObjAnuladoBln.ObjValorPro
            End If
        End If
    End Sub
    ''' <summary>
    ''' Determina si el valor ingresado es valido según las diferentes restricciones
    ''' </summary>
    ''' El valor pagado no es valido si el pago se esta haciendo a un agrupador de servicios habiendo 
    ''' deuda por pagar en otro agrupador de servicio o se esta haciendo un pago a un predio determinado
    ''' por un valor mayor al debido.
    ''' <remarks></remarks>
    Private Sub SValidePago()
        'Valor a pagar para el agrupador
        Dim ldecValorAnt = MobjPadre.ObjValorAnticipoDec.ObjValorPro
        ' Paga solo un servicio?
        Dim lblnPagaServicio = MobjPadre.ObjServicios_RecStr.ToString() <> "A"
        ' Es un anticipo
        Dim lblnEsAnticipo = MobjPadre.BlnEsSoloAnticipo
        Dim lstrIdpreAgr = MobjPadre.ObjIdPredioAgrupador_RecStr.ToString()
        Dim lstrServicios As String() = MobjPadre.ObjServicios_RecStr.ToString().Split(",")
        Dim ldecDeuda = MobjPadre.ObjClienteRecibo.FdecDeuda(lstrIdpreAgr.Split(","),
                lstrServicios)
        If lblnEsAnticipo Then
            If lblnPagaServicio Then
                If lstrServicios.Length > 1 OrElse Not lstrServicios.Contains("A") Then
                    HblnEsValido = GobjParametros.ObjPermiteAnticipoPorServicioBln.ObjValorPro
                End If
                If Not HblnEsValido Then
                    HstrMens = "No hay autorización para recibir Anticipos a un Servicio!"
                    SNotifiqueDatInv()
                Else
                    HblnEsValido = ldecDeuda = 0
                End If
            Else
                HblnEsValido = ldecDeuda = 0
            End If
        ElseIf lblnPagaServicio Then
            If MobjPadre.ObjValorAnticipoDec.ObjValorPro > 0 Then
                HblnEsValido = GobjParametros.ObjPermiteAnticipoPorServicioBln.ObjValorPro
                If Not HblnEsValido Then
                    HstrMens = "No hay autorización para recibir Anticipos a un Servicio!"
                    SNotifiqueDatInv()
                End If
            End If
            If HblnEsValido Then
                HblnEsValido = Not (HobjValorNew > ldecDeuda AndAlso
                        MobjPadre.ObjIdPredioAgrupador_RecStr.ToString().Contains(","))
                If Not HblnEsValido Then
                    HstrMens = "No es posible registrar el Anticipo, cuando se selecciona más de un " &
                            "Predio agrupador!"
                    SNotifiqueDatInv()
                End If
            End If
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando AndAlso HblnEsValido Then
            MobjPadre.ObjValorAnticipoDec.ObjValorPro = 0
            MobjPadre.ObjValorAnticipoDec.SValide()
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

Friend Class ClsValorAnticipoDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "ValorAnticipo"
    Private ReadOnly MobjPadre As ClsReciboCaja = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Valor Anticipo"
        HenuTipoValor = EnuTipoValor.EnuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0,
                Decimal.MaxValue, BlnEsRequerido, HenuTipoValor)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
                If Not MobjPadre.BlnEsSoloAnticipo Then
                    Dim ldecSaldoAPagar = MobjPadre.FdecSaldoRC()
                    Dim ldecPago As Decimal = MobjPadre.ObjValor_RecDec.ObjValorPro
                    If HobjValorNew <> 0 AndAlso ldecSaldoAPagar > 0 Then
                        HblnEsValido = False
                        HstrMens = "El Valor del Anticipo no es valido. El Recibo de Caja esta descuadrado!"
                        SNotifiqueDatInv()
                    End If
                End If
            Else
                If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando Then
                    If HobjValorNew = 0 Then
                        HblnEsValido = MobjPadre.ObjAnuladoBln.ObjValorPro
                    Else
                        HblnEsValido = HobjValorNew = HobjValorOriginal
                    End If
                Else
                    HblnEsValido = HobjValorNew = HobjValorOriginal
                End If
            End If
        Else
            HstrMens = "El Valor del Anticipo no es valido!"
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
            Return Format(HobjValorPro, "c")
        End If
    End Function
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
            MobjPadre.ObjValor_RecDec.SValide()
        End If
    End Sub
End Class

''' <summary>
''' Representa el valor debido al momento del pago sin tener en cuenta éste
''' </summary>
Friend Class ClsValorDeudaAlPagoDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "ValorDeudaAlPago"
    Private ReadOnly MobjPadre As ClsReciboCaja = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Valor Deuda al Pago"
        HenuTipoValor = EnuTipoValor.EnuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim ldecValorMinimo = 0.01D
        If MobjPadre.ObjValor_RecDec.ObjValorPro = MobjPadre.ObjValorAnticipoDec.ObjValorPro Then
            ldecValorMinimo = 0D
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, ldecValorMinimo,
                Decimal.MaxValue, BlnEsRequerido, HenuTipoValor)
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.EnuCreando Then
            HblnEsValido = HobjValorNew = HobjValorOriginal
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
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        MobjPadre.ObjSaldo_RecDec.SValide()
    End Sub
End Class
#End Region
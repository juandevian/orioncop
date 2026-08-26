Friend Class ClsAno
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriAnos"
    ' Enumerador
    ' Objetos de propiedad
    ' Colecciones
    Private McolPeriodos As Collection = Nothing
    Private MobjPeriodoActual As ClsPeriodo = Nothing
    Private McolServiciosAno As Collection = Nothing

    Private ReadOnly MobjPadre As ClsCentroUtilOriCop = GobjParametros
    Private MstrMensaje As String = String.Empty
#End Region

#Region "Constructores"
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Administrador al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwObjeto">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Friend Sub New(aobjPadre As ClsCentroUtilOriCop, adrwObjeto As DataRow)
        HobjPadre = aobjPadre
        MobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsAnulable = False
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
            Return EnuIdClasesPanDef.enuAno
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Año"
        End Get
    End Property
#End Region

#Region "Propiedades Prop"
    Friend ReadOnly Property ObjDiasParaDsctoPPShr As New ClsDiasParaDsctoPPShr(Me)
    Friend ReadOnly Property ObjDiasMultaExtShr As New ClsDiasMultaExtShr(Me)
    Friend ReadOnly Property ObjEstaCerradoAnoBln As New ClsEstaCerradoAnoBln(Me)
    Friend ReadOnly Property ObjFechaUltEstadoCtaDtm As New ClsFechaUltEstadoCtaDtm(Me)
    Friend ReadOnly Property ObjIdAnoShr As New ClsIdAnoShr(Me)
    Friend ReadOnly Property ObjIdCarpetaAnoShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtilAnoShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjModuloPorServicioBln As New ClsModuloPorServicioBln(Me)
    Friend ReadOnly Property ObjIdServicioMultaShr As New ClsIdServicioMultaShr(Me)
    Friend ReadOnly Property ObjTipoCalculoCuotaByt As New ClsTipoCalculoCuotaByt(Me)
    Friend ReadOnly Property ObjTipoDsctoPPByt As New ClsTipoDsctoPPByt(Me)
    Friend ReadOnly Property ObjTipoIncentivoByt As New ClsTipoIncentivoByt(Me)
    Friend ReadOnly Property ObjValorPres_AnoDec As New ClsValorPres_AnoDec(Me)
    Friend ReadOnly Property ObjValorMultaPagoExtDec As New ClsValorMultaPagoExtDec(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                With HcolPropiedades
                    .Add(ObjTipoIncentivoByt)
                    .Add(ObjDiasParaDsctoPPShr)
                    .Add(ObjDiasMultaExtShr)
                    .Add(ObjEstaCerradoAnoBln)
                    .Add(ObjFechaUltEstadoCtaDtm)
                    .Add(ObjIdAnoShr)
                    .Add(ObjIdCarpetaAnoShr)
                    .Add(ObjIdCentroUtilAnoShr)
                    .Add(ObjModuloPorServicioBln)
                    .Add(ObjIdServicioMultaShr)
                    .Add(ObjTipoCalculoCuotaByt)
                    .Add(ObjTipoDsctoPPByt)
                    .Add(ObjValorMultaPagoExtDec)
                    .Add(ObjValorPres_AnoDec)
                End With
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region

#Region "Otras propiedades"
    Friend Property BlnInicioMesActual As Boolean = Nothing
    Friend ReadOnly Property DtmFechaInicioAno As Date
        Get
            Return DateSerial(ObjIdAnoShr.ObjValorPro, 1, 1)
        End Get
    End Property

    Friend ReadOnly Property DtmFechaFinAno As Date
        Get
            Return DateSerial(ObjIdAnoShr.ObjValorPro, 12, 31)
        End Get
    End Property

    Friend ReadOnly Property ObjPeriodoActual As ClsPeriodo
        Get
            Dim lblnNoEstaCerrado = False
            For Each lobjPerio As ClsPeriodo In ColPeriodos
                If Not lobjPerio.ObjEstaCerradoPeriodoBln.ObjValorPro Then
                    MobjPeriodoActual = lobjPerio
                    lblnNoEstaCerrado = True
                    Exit For
                End If
            Next
            If Not lblnNoEstaCerrado AndAlso ColPeriodos.Count = 12 Then
                MobjPeriodoActual = McolPeriodos(ColPeriodos.Count)
            End If
            Return MobjPeriodoActual
        End Get
    End Property

    ''' <summary>
    ''' Devuelve el nombre del mes actual seguido del año actual
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend ReadOnly Property StrNombrePeriodoActual As String
        Get
            Dim lstrNombre = String.Empty
            Dim lobjPerAct As ClsPeriodo = ObjPeriodoActual
            If Not IsNothing(lobjPerAct) Then
                lstrNombre = FstrMesFecha(lobjPerAct.DtmFechaFinPeriodo) & " " &
                        ObjIdAnoShr.ToString
            End If
            Return lstrNombre
        End Get
    End Property

    ''' <summary>
    ''' Devuelve una cadena de seis caracteres formada por los cuatro caracteres del año actual y los dos
    ''' caracteres del mes actual. (yyyymm)
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend ReadOnly Property StrIdPeriodoActual As String
        Get
            Dim lstrIdPeriodo = ObjPeriodoActual.ObjIdAnoPeriodoShr.ToString() &
                    ObjPeriodoActual.ObjIdPeriodoShr.ToString()
            Return lstrIdPeriodo
        End Get
    End Property
#End Region

#End Region

#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SVacie()
        MyBase.SVacie()
        MobjPeriodoActual = Nothing
        McolPeriodos = Nothing
        McolServiciosAno = Nothing
    End Sub
    Protected Overrides Sub SInicialiceObj()
        ObjIdCarpetaAnoShr.ObjValorPro = GshrIdCarpeta
        ObjIdCentroUtilAnoShr.ObjValorPro = GshrIdCentroUtil
        ObjTipoIncentivoByt.ObjValorPro = 0
        ObjDiasParaDsctoPPShr.ObjValorPro = 0
        ObjEstaCerradoAnoBln.ObjValorPro = False
        ObjModuloPorServicioBln.ObjValorPro = False
        ObjTipoDsctoPPByt.ObjValorPro = 0
        ObjValorPres_AnoDec.ObjValorPro = 0
        ObjFechaUltEstadoCtaDtm.ObjValorPro = GCDTMFECHANULA
    End Sub
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        Dim lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            GobjPanDat.SInicialiceTransaccion()
            If EnuEstadoActualizacion <> EnuEstadoObjetoDef.EnuConsultando Then
                If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
                    SGenerePeriodos()
                Else
                    SSincroniceServicios()
                End If
                SActualiceDsctoPPSect()
                SEstablezcaAjustar()
                ClsPanorama.SActualiceCol(ColPeriodos)
                ClsPanorama.SActualiceCol(ColServiciosAno)
                MyBase.SActualice(ablnExigeRequeridos)
                MobjPadre.SRefresqueObj()
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
            If Not lblnNoHayError Then
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            Else
                GobjPanDat.SConfirmeTransaccion()
                GobjPanDat.SControleProcesoObj(False)
            End If
        End Try
    End Sub
    Friend Overrides Function FblnEsModificable() As Boolean
        Dim lblnEsModi = MyBase.FblnEsModificable()
        If lblnEsModi Then
            lblnEsModi = Not ObjEstaCerradoAnoBln.ObjValorPro
        End If
        Return lblnEsModi
    End Function
    Friend Sub SSuprima(ByRef astrMens As String)
        MstrMensaje = String.Empty
        EnuPermisosObj += EnuPermisosDef.EnuSuprimir
        If FblnEsSuprimible() Then
            If Not FblnSuprimio() Then
                MstrMensaje = "No fue posible suprimir el año!"
            End If
        End If
        astrMens = MstrMensaje
        MstrMensaje = String.Empty
    End Sub
    Protected Overrides Function FblnSuprimio() As Boolean
        Dim lblnSuprimio = False, lblnNoHayError As Boolean
        Try
            GobjPanDat.SControleProcesoObj(True)
            EnuPermisosObj += EnuPermisosDef.EnuSuprimir
            If FblnPermitidoSuprimir() Then
                GobjPanDat.SInicialiceTransaccion()
                lblnSuprimio = ClsPanorama.FblnSuprimioCol(ColPeriodos)
                If lblnSuprimio Then
                    For Each lobjServicios As ClsServicio In ColServiciosAno
                        lobjServicios.EnuPermisosObj += EnuPermisosDef.EnuSuprimir
                    Next
                    lblnSuprimio = ClsPanorama.FblnSuprimioCol(ColServiciosAno)
                End If
                If lblnSuprimio Then
                    SSuprimaItemsProFac()
                    lblnSuprimio = MyBase.FblnSuprimio()
                End If
                If lblnSuprimio Then
                    GobjPanorama.SRegistreAccionLogApp(HstrNombreClase, "Suprimir Año " &
                            ObjIdAnoShr.ToString)
                    GobjPanDat.SConfirmeTransaccion()
                Else
                    GobjPanDat.SAborteTransaccion()
                End If
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
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
        Return lblnSuprimio
    End Function
    Friend Overrides Function FblnEsSuprimible() As Boolean
        Dim lblnEsSuprimible = FblnPermitidoSuprimir()
        If lblnEsSuprimible Then
            Dim lshrIdAnoActual As Short = GobjParametros.ObjAnoActual.ObjIdAnoShr.ObjValorPro
            lblnEsSuprimible = (ObjIdAnoShr.ObjValorPro > lshrIdAnoActual)
            If lblnEsSuprimible Then
                lblnEsSuprimible = GobjParametros.FblnEsElUltimoAno(ObjIdAnoShr.ObjValorPro)
                If Not lblnEsSuprimible Then
                    MstrMensaje = "El único año suprimible es el último diferente al año actual!"
                End If
            Else
                MstrMensaje = "El Año actual no puede ser eliminado!"
            End If
        Else
            MstrMensaje = "El Usuario actual no tiene permiso para está acción!"
        End If
        Return lblnEsSuprimible
    End Function
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjIdAnoShr.ToString
        End Get
    End Property
    Protected Overrides Sub SCreeObj(aobjValorLlave() As Object)
        Dim lblnNoHayError = False, lblnEsCreable = False
        Try
            lblnEsCreable = FblnEsCreable(aobjValorLlave)
            If lblnEsCreable Then
                EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando
                SVacie()
                SInicialiceObj()
                ObjValorLlave = aobjValorLlave
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
                If Not lblnEsCreable Then
                    Throw New PanLException("Año no creable!")
                End If
            End If
        End Try
    End Sub
#End Region

#Region "Procedimientos del objeto"
    Friend Function FblnEsAnoActual() As Boolean
        Dim lshrIdAnoActual = 0S
        If GobjParametros.ObjAnoActual IsNot Nothing Then
            lshrIdAnoActual = GobjParametros.ObjAnoActual.ObjIdAnoShr.ObjValorPro
        End If
        Return lshrIdAnoActual = ObjIdAnoShr.ObjValorPro
    End Function

    Friend Function FdecValorPresupuesto() As Decimal
        Dim ldecPres = 0D
        If ColServiciosAno.Count > 0 Then
            For Each lobjServicio As ClsServicio In ColServiciosAno
                With lobjServicio
                    If .BlnEsCuotaAdministracion Then
                        ldecPres += .FdecValor
                    End If
                End With
            Next
        End If
        Return ldecPres
    End Function

    ''' <summary>
    ''' Devuelve el valor real a cobrar en el año por la Cuota de Administración.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>Se calcula con base en la programación de facturas.</remarks>
    Friend Function FdecValorCuotaAdminAno() As Decimal
        Dim ldecVlrTot = 0D
        For Each lobjServicio As ClsServicio In ColServiciosAno
            If lobjServicio.BlnEsCuotaAdministracion AndAlso
                    Not lobjServicio.ObjEsAjusteBln.ObjValorPro Then
                ldecVlrTot += ClsOrionCop.FdecValorTotalCalculadoServicio(
                        lobjServicio.ObjIdAno_ServicioShr.ObjValorPro,
                        lobjServicio.ObjIdServicioShr.ObjValorPro)
            End If
        Next
        Return ldecVlrTot
    End Function

    Friend Function FblnDebeImportarAjuste() As Boolean
        Dim lblnDebeImporAjuste = False
        If ObjTipoCalculoCuotaByt.ObjValorPro = EnuTipoBaseCalculo.EnuImportadas Then
            Dim lstrUltimoPeriodoIni = String.Empty
            For Each lobjSer As ClsServicio In ColServiciosAno
                If lobjSer.ObjEsAjusteBln.ObjValorPro Then
                    If lobjSer.ObjPeriodoInicioStr.ToString() > lstrUltimoPeriodoIni Then
                        lstrUltimoPeriodoIni = lobjSer.ObjPeriodoInicioStr.ToString()
                    End If
                End If
            Next
            For Each lobjSer As ClsServicio In ColServiciosAno
                If lobjSer.ObjEsAjusteBln.ObjValorPro Then
                    If lobjSer.ObjPeriodoInicioStr.ToString() = lstrUltimoPeriodoIni Then
                        Dim lshrIdAno As Short = ObjIdAnoShr.ObjValorPro
                        Dim lshrIdSerAjustado As Short = lobjSer.ObjIdServicioAjustadoShr.ObjValorPro
                        Dim lobjSerAjustado As ClsServicio = ColServiciosAno(lshrIdAno.ToString &
                                "," & lshrIdSerAjustado)
                        lblnDebeImporAjuste = lobjSer.DecValorAjuste = 0 AndAlso
                                Not lobjSerAjustado.ObjEstaAjustadoBln.ObjValorPro
                        If lblnDebeImporAjuste Then Exit For
                    End If
                End If
            Next
        End If
        Return lblnDebeImporAjuste
    End Function

    Friend Function FblnFacturacionGenerada() As Boolean
        Dim lblnFacturado = False
        If ColPeriodos IsNot Nothing Then
            lblnFacturado = ObjPeriodoActual.ObjFechaFacturacionPeriodoDtm.ObjValorPro <>
                    GCDTMFECHANULA
        End If
        Return lblnFacturado
    End Function

    Friend Function FstrUltimoPerFacturado() As String
        Dim lstrUltPerFac = String.Empty
        Dim lobjPeriodo As ClsPeriodo
        If ColPeriodos.Count > 0 Then
            For i As Integer = 12 To 1 Step -1
                lobjPeriodo = McolPeriodos(i)
                If lobjPeriodo.ObjEstaCerradoPeriodoBln.ObjValorPro Then
                    lstrUltPerFac = lobjPeriodo.StrPeriodo
                    Exit For
                Else
                    If i = 12 Then
                        If GobjParametros.FblnEsElPrimerAno(ObjIdAnoShr.ObjValorPro) Then
                            If GobjParametros.FblnPerActEsDicPrimerAno Then
                                lstrUltPerFac = lobjPeriodo.StrPeriodo
                                Exit For
                            End If
                        End If
                    End If
                    If lobjPeriodo.ObjFechaFacturacionPeriodoDtm.ObjValorPro <> GCDTMFECHANULA Then
                        lstrUltPerFac = lobjPeriodo.StrPeriodo
                        Exit For
                    End If
                End If
            Next
        End If
        Return lstrUltPerFac
    End Function

    Friend Function FblnEstaGenCuota() As Boolean
        Dim lblnEstaGen = False
        If ColServiciosAno.Count > 0 Then
            For Each lobjSer As ClsServicio In ColServiciosAno
                If Not lobjSer.ObjEsAjusteBln.ObjValorPro Then
                    lblnEstaGen = lobjSer.ObjEstaGenaradaProgramBln.ObjValorPro
                End If
                If Not lblnEstaGen Then Exit For
            Next
        End If
        Return lblnEstaGen
    End Function

    Friend Function FblnValorPresExigeCero() As Boolean
        Dim lblnExige = False
        Dim lobjPrimerAno As ClsAno = GobjParametros.ColAnos(1)
        Dim lblnEsPrimerAno = (lobjPrimerAno.ObjIdAnoShr.ObjValorPro = ObjIdAnoShr.ObjValorPro)
        If lblnEsPrimerAno Then
            lblnExige = GobjParametros.FblnPerActEsDicPrimerAno
        End If
        Return lblnExige
    End Function

    ''' <summary>
    ''' Determina si el cálculo de la cuota de administración se hace con base en el coeficiente
    ''' de propiedad o con base en el porcentaje de participacón de los sectores; se hace con base 
    ''' en el coeficiente de propiedad cuando todos los sectores participan con el 100%, de lo
    ''' contrario se hace con base en las áreas y los porcentajes de participación de los sectores
    ''' </summary>
    ''' <returns></returns>
    Friend Function FblnCalcularCuotaAdminPorCP() As Boolean
        Dim lblnSi As Boolean
        For Each lobjServicio As ClsServicio In ColServiciosAno
            If Not lobjServicio.ObjEsAjusteBln.ObjValorPro Then
                lblnSi = lobjServicio.FblnSectContrConTotalArea
                If Not lblnSi Then
                    Exit For
                End If
            End If
        Next
        Return lblnSi
    End Function

    Friend Sub SActualiceFechaGenEstados(adtmFechaUltGenEtados As Date)
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
            EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
        End If
        ObjFechaUltEstadoCtaDtm.ObjValorPro = adtmFechaUltGenEtados
        SActualice(True)
    End Sub

    Friend Sub SVerifiqueApp()
        Dim lstrMens = GobjParametros.SVerifiqueApp(False, True)
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEventoNot(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        Else
            SLevanteEventoNot(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuOk)
        End If
    End Sub

    Friend Function FblnAnoEsModificable(ByRef astrMens As String) As Boolean
        Dim lblnEsModi = FblnEsAnoActual() OrElse ObjIdAnoShr.ObjValorPro >
                GobjParametros.ObjAnoActual.ObjIdAnoShr.ObjValorPro
        If lblnEsModi Then
            lblnEsModi = ObjPeriodoActual.ObjIdPeriodoShr.ObjValorPro < 12
            If Not lblnEsModi Then
                lblnEsModi = Not ObjPeriodoActual.BlnPeriodoFacturado
            End If
            If Not lblnEsModi Then
                astrMens = "El Año no puede ser modificado cuando diciembre ya está facturado!"
            End If
        Else
            astrMens = "Solo puede ser modificado el año actual!"
        End If
        Return lblnEsModi
    End Function

    Friend Function FobjServicioMulta() As ClsServicio
        Dim lobjSerMulta As ClsServicio = Nothing
        If GobjParametros.ColServiciosPer.Count > 0 Then
            Dim lstrKey = "0," & ObjIdServicioMultaShr.ObjValorPro.ToString
            If GobjParametros.ColServiciosPer.Contains(lstrKey) Then
                lobjSerMulta = GobjParametros.ColServiciosPer(lstrKey)
            End If
        End If
        Return lobjSerMulta
    End Function
#End Region

#Region "Manejo Periodos"
    Friend ReadOnly Property ColPeriodos As Collection
        Get
            If ObjIdAnoShr.BlnEsValido AndAlso EnuEstadoActualizacion <> EnuEstadoObjetoDef.EnuCreando Then
                If IsNothing(McolPeriodos) Then
                    McolPeriodos = New Collection
                    Dim ldtbPeriodos = FdtbPeriodos()
                    For Each ldrwPeriodo As DataRow In ldtbPeriodos.Rows
                        Dim lobjPeriodo As New ClsPeriodo(Me, ldrwPeriodo)
                        lobjPeriodo.SLeaValores(True)
                        McolPeriodos.Add(lobjPeriodo, lobjPeriodo.ObjIdPeriodoShr.ToString)
                    Next
                End If
            End If
            Return McolPeriodos
        End Get
    End Property
    Friend ReadOnly Property DtbPeriodos
        Get
            Return FdtbPeriodos()
        End Get
    End Property
    Private Sub SSuprimaItemsProFac()
        Dim lstrTabla = ClsItemProgramaFact.SstrNombreTabla
        Dim lcolCamposRef As New Collection
        Dim lcolDatosRef As New Collection
        lcolCamposRef.Add(StrCampoCarpeta)
        lcolCamposRef.Add(StrCampoCentroUtil)
        lcolCamposRef.Add(ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd)
        lcolDatosRef.Add(GshrIdCarpeta, StrCampoCarpeta)
        lcolDatosRef.Add(GshrIdCentroUtil, StrCampoCentroUtil)
        lcolDatosRef.Add(ObjIdAnoShr.ObjValorPro, ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd)
        GobjPanDat.SElimineRegistro(lstrTabla, lcolCamposRef, lcolDatosRef)
    End Sub
    Private Sub SGenerePeriodos()
        Dim ldtbPeriodos = FdtbPeriodos()
        Dim ldrwPeriodo As DataRow
        Dim lblnEsPrimerAno = (MobjPadre.ColAnos.Count = 0)
        If ldtbPeriodos.Rows.Count > 0 Then
            Throw New ErrorInesperadoPanLException("DataTable con registros en clsAno.sGenerePeriodos")
        End If
        McolPeriodos = New Collection
        For i As Short = 1 To 12
            ldrwPeriodo = ldtbPeriodos.NewRow
            Dim lobjPeriodo As New ClsPeriodo(Me, ldrwPeriodo)
            lobjPeriodo.EnuPermisosObj += EnuPermisosDef.EnuCrear
            With lobjPeriodo
                .SCreeObj(Nothing)
                .ObjIdCarpetaPeriodoShr.ObjValorPro = GshrIdCarpeta
                .ObjIdCentroUtilPeriodoShr.ObjValorPro = GshrIdCentroUtil
                .ObjIdAnoPeriodoShr.ObjValorPro = ObjIdAnoShr.ObjValorPro
                .ObjIdPeriodoShr.ObjValorPro = i
                .ObjEstaCerradoPeriodoBln.ObjValorPro = False
                .ObjFechaFacturacionPeriodoDtm.ObjValorPro = GCDTMFECHANULA
            End With
            McolPeriodos.Add(lobjPeriodo, lobjPeriodo.ObjIdPeriodoShr.ToString)
        Next
        If lblnEsPrimerAno Then
            SCierrePeriodos()
        End If
    End Sub
    Private Sub SSincroniceServicios()
        If Not GblnActualizandoApp Then
            If ObjTipoCalculoCuotaByt.BlnCambio Then
                For Each lobjServicio As ClsServicio In ColServiciosAno
                    If Not lobjServicio.ObjEsAjusteBln.ObjValorPro Then
                        With lobjServicio
                            If lobjServicio.EnuEstadoActualizacion =
                                    EnuEstadoObjetoDef.EnuConsultando Then
                                .EnuEstadoActualizacion =
                                        EnuEstadoObjetoDef.EnuModificando
                            End If
                            .ObjTipoBaseCalculoByt.ObjValorPro =
                                    ObjTipoCalculoCuotaByt.ObjValorPro
                        End With
                    End If
                Next
            End If
            If ObjValorPres_AnoDec.BlnCambio AndAlso ObjModuloPorServicioBln.ObjValorPro AndAlso
                        Not GblnImportando Then
                Dim lshrIdAno As Short = ObjIdAnoShr.ObjValorPro
                Dim lblCtasYaGeneradas = ClsOrionCop.FblnEstanCalcuCuotasAdmin(lshrIdAno)
                For Each lobjServicio As ClsServicio In ColServiciosAno
                    If Not lobjServicio.ObjEsAjusteBln.ObjValorPro Then
                        With lobjServicio
                            If Not ClsOrionCop.BlnProcesoEspecial Then
                                If lobjServicio.EnuEstadoActualizacion =
                                        EnuEstadoObjetoDef.EnuConsultando Then
                                    .EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
                                End If
                                .ObjEstaGenaradaProgramBln.ObjValorPro = False
                                .ObjEstaAjustadoBln.ObjValorPro = Not lblCtasYaGeneradas
                                .SActualice(True)
                            End If
                        End With
                    End If
                Next
            End If
        End If
    End Sub
    Private Sub SActualiceDsctoPPSect()
        If Not ObjTipoIncentivoByt.ObjValorPro = EnuTipoIncentivo.EnuDescuentoPP Then
            For Each lobjSector As ClsSector In GobjParametros.ColSectores
                If lobjSector.ObjDctoProntoPago_SecDbl.ObjValorPro > 0 Then
                    lobjSector.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
                    lobjSector.ObjDctoProntoPago_SecDbl.ObjValorPro = 0
                    lobjSector.SActualice(True)
                End If
            Next
        End If
    End Sub
    Private Function FdtbPeriodos() As DataTable
        Dim lstrIndice = {{StrCampoCarpeta, "ASC"},
                          {StrCampoCentroUtil, "ASC"},
                          {ClsIdAnoShr.SstrNombreCampoBd, "ASC"},
                          {ClsIdPeriodoShr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsIdAnoShr.SstrNombreCampoBd &
                    " = " & ObjIdAnoShr.ToString
        Dim lstrCamposSelect() = {"*", "'Nombre' AS Nombre"}
        Dim ldtbPeriodos = ClsPanorama.FdtbDataTable(ClsPeriodo.SstrNombreTabla, lstrCamposSelect,
                    lstrIndice, lstrFiltro)
        If ldtbPeriodos.Rows.Count > 0 Then
            Dim ldrwMes As DataRow
            For i = 0 To ldtbPeriodos.Rows.Count - 1
                ldrwMes = ldtbPeriodos.Rows(i)
                ldrwMes("Nombre") = MonthName(i + 1)
            Next
        End If
        Return ldtbPeriodos
    End Function
    Private Sub SCierrePeriodos()
        Dim lobjPeriodo As ClsPeriodo
        Dim lentId1erPeriodoAbierto = 0
        If Date.Today.Month = 1 Then
            If Not BlnInicioMesActual Then
                lentId1erPeriodoAbierto = 11
            Else
                lentId1erPeriodoAbierto = 12
            End If
        ElseIf Date.Today.Month = 2 Then
            If Not BlnInicioMesActual Then
                lentId1erPeriodoAbierto = 12
            Else
                lentId1erPeriodoAbierto = 1
            End If
        ElseIf Date.Today.Month > 2 Then
            If Not BlnInicioMesActual Then
                lentId1erPeriodoAbierto = Date.Today.Month - 2
            Else
                lentId1erPeriodoAbierto = Date.Today.Month - 1
            End If
        End If
        Dim lshrIdAno As Short = ObjIdAnoShr.ObjValorPro
        For lentIdPeriodo As Integer = 1 To lentId1erPeriodoAbierto
            lobjPeriodo = McolPeriodos(lentIdPeriodo)
            If lentIdPeriodo < lentId1erPeriodoAbierto Then
                lobjPeriodo.ObjEstaCerradoPeriodoBln.ObjValorPro = True
            End If
            If lentIdPeriodo <= lentId1erPeriodoAbierto Then
                lobjPeriodo.ObjFechaFacturacionPeriodoDtm.ObjValorPro =
                        DateSerial(lshrIdAno, lentIdPeriodo, 1)
            End If
        Next
    End Sub
    Friend Sub SCierrePeriodoActual()
        With ObjPeriodoActual
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
                .EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
            End If
            .ObjEstaCerradoPeriodoBln.ObjValorPro = True
            .SActualice(True)
        End With
    End Sub
#End Region

#Region "Manejo Servicios Ano"
    Friend ReadOnly Property ColServiciosAno As Collection
        Get
            If McolServiciosAno Is Nothing OrElse McolServiciosAno.Count = 0 Then
                Dim ldtbServiciosAno = FdtbServiciosAno()
                McolServiciosAno = New Collection
                For Each ldrwServicioAno As DataRow In ldtbServiciosAno.Rows
                    Dim lobjServicioAno As New ClsServicio(Me, ldrwServicioAno)
                    lobjServicioAno.SLeaValores(True)
                    Dim lstrKey = lobjServicioAno.ObjIdAno_ServicioShr.ToString & "," &
                        lobjServicioAno.ObjIdServicioShr.ToString
                    McolServiciosAno.Add(lobjServicioAno, lstrKey)
                Next
            End If
            Return McolServiciosAno
        End Get
    End Property
    Friend Sub SReleaColServicios()
        McolServiciosAno = Nothing
        McolServiciosAno = ColServiciosAno
        GobjParametros.SRefresqueObj()
    End Sub
    Friend Sub SRemuevaServicio(astrKey As String)
        If ColServiciosAno.Contains(astrKey) Then
            McolServiciosAno.Remove(astrKey)
        End If
    End Sub
    Friend Function FobjNuevoServicioAno(ablnAjuste As Boolean) As ClsServicio
        Dim ldtbServiciosAno = FdtbServiciosAno()
        Dim lstrNombreServicio = String.Empty
        Dim lblnModificoPermiso = False
        Dim ldrwServicioNuevo As DataRow = ldtbServiciosAno.NewRow
        If ColServiciosAno.Count = 0 Then
            lstrNombreServicio = "Cuota Administración"
        End If
        Dim lobjNuevoServicioAno = New ClsServicio(Me, ldrwServicioNuevo)
        With lobjNuevoServicioAno
            If ablnAjuste AndAlso Not CType(.EnuPermisosObj And EnuPermisosDef.EnuCrear, Boolean) Then
                .EnuPermisosObj += EnuPermisosDef.EnuCrear
                lblnModificoPermiso = True
            End If
            .SCreeObj(Nothing)
            .BlnCreandoAno = True
            If lblnModificoPermiso Then
                .EnuPermisosObj -= EnuPermisosDef.EnuCrear
            End If
        End With
        Return lobjNuevoServicioAno
    End Function
    Friend Function FblnExisteNombreServicio(astrNombreServicio As String) As Boolean
        Dim ldtbServiciosAno = FdtbServiciosAno()
        Dim lstrFiltro = ClsNombreServicioStr.SstrNombreCampoBd & " = '" & astrNombreServicio & "'"
        Dim ldrwServicios() As DataRow = ldtbServiciosAno.Select(lstrFiltro)
        Return (ldrwServicios.Length > 0)
    End Function
    Private Function FdtbServiciosAno() As DataTable
        Dim lshrIdAno As Short = 9999, ldtbSerAno As DataTable
        If Not IsNothing(ObjIdAnoShr.ObjValorPro) Then
            lshrIdAno = ObjIdAnoShr.ObjValorPro
        End If
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdAnoShr.SstrNombreCampoBd & " = " & lshrIdAno
        Dim lstrOrden As String(,) = {{ClsIdServicioShr.SstrNombreCampoBd, "ASC"}}
        ldtbSerAno = ClsPanorama.FdtbDataTable(ClsServicio.SstrNombreTabla, {"*"},
                    lstrOrden, lstrFiltro)
        Return ldtbSerAno
    End Function
    Friend Function FobjServicioAjuste(aobjServicioAAjustar As ClsServicio,
            aentCantPeriodos As Integer) As ClsServicio
        Dim lobjServicioAjuste As ClsServicio = FobjNuevoServicioAno(True)
        Dim lshrIdServicio As Short = ColServiciosAno.Count + 1
        With lobjServicioAjuste
            .ObjIdTipoServicioByt.ObjValorPro = EnuTipoServicio.EnuAnual
            .ObjIdAno_ServicioShr.ObjValorPro = ObjIdAnoShr.ObjValorPro
            .ObjIdServicioShr.ObjValorPro = lshrIdServicio
            .ObjEsAjusteBln.ObjValorPro = True
            .ObjEstaAjustadoBln.ObjValorPro = False
            .ObjIdServicioAjustadoShr.ObjValorPro = aobjServicioAAjustar.ObjIdServicioShr.ObjValorPro
            .ObjDiaFacturaShr.ObjValorPro = aobjServicioAAjustar.ObjDiaFacturaShr.ObjValorPro
            .ObjDiasGraciaShr.ObjValorPro = aobjServicioAAjustar.ObjDiasGraciaShr.ObjValorPro
            .ObjDiasVencimientoShr.ObjValorPro = aobjServicioAAjustar.ObjDiasVencimientoShr.ObjValorPro
            .ObjVenceFinMesBln.ObjValorPro = aobjServicioAAjustar.ObjVenceFinMesBln.ObjValorPro
            .ObjFactAPropYPreAgrBln.ObjValorPro = False
            .ObjEsFactProgramableBln.ObjValorPro = True
            .ObjGeneraProgramBln.ObjValorPro = True
            .ObjNombreServicioStr.ObjValorPro = "Retroactivo " & aobjServicioAAjustar.ObjNombreServicioStr.ObjValorPro
            .ObjConceptoServicioStr.ObjValorPro = "Retroactivo " & aobjServicioAAjustar.ObjNombreServicioStr.ObjValorPro
            .ObjEstaActivoServicioBln.ObjValorPro = False
            .ObjBaseMinimaReteFuenteDec.ObjValorPro = aobjServicioAAjustar.ObjBaseMinimaReteFuenteDec.ObjValorPro
            .ObjCantPeriodos_ServicioShr.ObjValorPro = aentCantPeriodos
            .ObjModoCausaInteresesByt.ObjValorPro =
                    aobjServicioAAjustar.ObjModoCausaInteresesByt.ObjValorPro
            .ObjCodigoCuentaCrStr.ObjValorPro = aobjServicioAAjustar.ObjCodigoCuentaCrStr.ObjValorPro
            .ObjCodigoCuentaDbStr.ObjValorPro = aobjServicioAAjustar.ObjCodigoCuentaDbStr.ObjValorPro
            .ObjCodigoCuentaIvaStr.ObjValorPro = aobjServicioAAjustar.ObjCodigoCuentaIvaStr.ObjValorPro
            .ObjCodigoCuentaDevStr.ObjValorPro = aobjServicioAAjustar.ObjCodigoCuentaDevStr.ObjValorPro
            .ObjCodigoCuentaMoraStr.ObjValorPro = aobjServicioAAjustar.ObjCodigoCuentaMoraStr.ObjValorPro
            .ObjEsExcluidoIvaBln.ObjValorPro = aobjServicioAAjustar.ObjEsExcluidoIvaBln.ObjValorPro
            .ObjEstaGenaradaProgramBln.ObjValorPro = True
            .ObjEsServicioIdBln.ObjValorPro = False
            .ObjTipoBaseCalculoByt.ObjValorPro = aobjServicioAAjustar.ObjTipoBaseCalculoByt.ObjValorPro
            .ObjIdTipoTerCtaCrSerByt.ObjValorPro = aobjServicioAAjustar.ObjIdTipoTerCtaCrSerByt.ObjValorPro
            .ObjPeriodoInicioStr.ObjValorPro = StrIdPeriodoActual
            .ObjTarifaIvaDbl.ObjValorPro = aobjServicioAAjustar.ObjTarifaIvaDbl.ObjValorPro
            .ObjTarifaRetFteDbl.ObjValorPro = aobjServicioAAjustar.ObjTarifaRetFteDbl.ObjValorPro
            .ObjBaseMinimaReteIcaDec.ObjValorPro = 0
            .ObjTarifaRetIcaDbl.ObjValorPro = 0
            .SActualice(True)
        End With
        SReleaColServicios()
        Dim lstrKey = ObjIdAnoShr.ToString() & "," & lshrIdServicio.ToString()
        lobjServicioAjuste = ColServiciosAno(lstrKey)
        Return lobjServicioAjuste
    End Function
    ''' <summary>
    ''' Ajusta en el año el presupuesto con base en los valores de los servicios cuota 
    ''' de administración.
    ''' </summary>
    ''' <remarks>Siempre que hay un cambio en el valor de un servicio se debe ajustar aqui en el año</remarks>
    Friend Sub SAjustePresupuesto()
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando AndAlso
                Not ObjModuloPorServicioBln.ObjValorPro Then
            SRefresqueObj()
            Dim ldecPreSer = FdecValorPresupuesto()
            If ldecPreSer <> ObjValorPres_AnoDec.ObjValorPro Then
                EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
                ObjValorPres_AnoDec.ObjValorPro = ldecPreSer
                SActualice(True)
            End If
        End If
    End Sub
    ''' <summary>
    ''' Si el presupuesto del año cambió y ya se a llevado a cabo la facturación en algún 
    ''' período se debe ajustar el cobro por administración
    ''' </summary>
    Private Sub SEstablezcaAjustar()
        If ObjValorPres_AnoDec.BlnCambio Then
            Dim lblnHayFactAno = FblnHayFacsEnAno()
            For Each lobjSer As ClsServicio In ColServiciosAno
                If Not lobjSer.ObjEsAjusteBln.ObjValorPro Then
                    If lobjSer.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
                        If Not lobjSer.EnuPermisosObj And EnuPermisosDef.EnuModificar Then
                            lobjSer.EnuPermisosObj += EnuPermisosDef.EnuModificar
                        End If
                        lobjSer.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
                    ElseIf lobjSer.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando Then
                        If Not lblnHayFactAno OrElse FdecValorCuotaAdminAno() = 0 Then
                            lobjSer.ObjEstaAjustadoBln.ObjValorPro = True
                        ElseIf Not GblnImportando Then
                            lobjSer.ObjEstaAjustadoBln.ObjValorPro = False
                        End If
                    Else
                        lobjSer.ObjEstaAjustadoBln.ObjValorPro = True
                    End If
                End If
            Next
        End If
    End Sub
    Friend Function FblnHayFacsEnAno() As Boolean
        Dim lstrTabla = ClsFactura.SstrNombreTabla
        Dim lstrCampSele = {"COUNT(" & ClsIdFacturaEnt.SstrNombreCampoBd & ")"}
        Dim lstrOrden = {{"", ""}}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND YEAR(" &
                ClsFechaFacturaDtm.SstrNombreCampoBd & ") = " & ObjIdAnoShr.ObjValorPro
        Dim ldtbFacs = ClsPanorama.FdtbDataTable(lstrTabla, lstrCampSele, lstrOrden, lstrFiltro)
        Dim lblnHay = ClsPanorama.FobjValorCampo(ldtbFacs.Rows(0)(0), EnuTipoValor.EnuInteger) > 0
        Return lblnHay
    End Function
    Friend Function FblnModuloYaContribuye(ashrIdServicio As Short) As Boolean
        Dim lblnYaCont = False
        For Each lobjSer As ClsServicio In ColServiciosAno
            lblnYaCont = lobjSer.FblnModuloMeContribuye(ashrIdServicio)
            If lblnYaCont Then Exit For
        Next
        Return lblnYaCont
    End Function
#End Region

#Region "Estado del Proceso de Calculo"
    Friend Function FblnCalculoCuotasAdmin(ByRef astrMens As String)
        Dim lblnCalculo = False
        If FblnCalcularCuotas(astrMens) Then
            Dim lobjCalSer As New ClsCalculosServicios
            lblnCalculo = lobjCalSer.FblnCalculoCuotasAdmin(Me, astrMens)
            If lblnCalculo Then
                SRegisCalculo()
            End If
        End If
        Return lblnCalculo
    End Function
    ''' <summary>
    ''' Indica si se pueden calcular las cuotas de administración. El programa determina si
    ''' calcula las iniciales o las definitivas
    ''' </summary>
    ''' <returns></returns>
    Friend Function FblnCalcularCuotas(ByRef astrMens As String) As Boolean
        Dim lshrIdAnoActual = GobjParametros.ObjAnoActual.ObjIdAnoShr.ObjValorPro
        Dim lblnPuedeCalCuotas = ObjIdAnoShr.ObjValorPro >= lshrIdAnoActual
        If lblnPuedeCalCuotas Then
            Dim lstrMens = String.Empty
            lblnPuedeCalCuotas = GobjParametros.FblnParaCuotaAdminOk(lstrMens)
            If lblnPuedeCalCuotas Then
                lblnPuedeCalCuotas = FblnPresOk(astrMens)
            End If
            If Not lblnPuedeCalCuotas Then
                If astrMens.Length = 0 Then
                    astrMens = "Hay problemas con la parametrización del presupuesto o de las cuotas!"
                End If
            Else
                lblnPuedeCalCuotas = ObjTipoCalculoCuotaByt.ObjValorPro <>
                        EnuTipoBaseCalculo.EnuImportadas
                If Not lblnPuedeCalCuotas Then
                    astrMens = "No es posible calcular las cuotas cuando estas han sido importadas!"
                End If
            End If
        Else
            astrMens = "Solo se pueden calcular Cuotas de Administración del año actual o posterior!"
        End If
        Return lblnPuedeCalCuotas
    End Function
    ''' <summary>
    ''' Indica si las cuotas basadas en el presupuesto inicial pueden ser calculadas
    ''' </summary>
    ''' <remarks>El calculo de las cuotas de administración se basan en elpresupuesto inicial
    ''' o en el presupuesto definitivo. No puede existir la posibilidad de hacer el calculo
    ''' basado en los dos presupuestos</remarks>
    ''' <returns></returns>
    Private Function FblnPresOk(ByRef astrMens As String) As Boolean
        Dim lblnPresOk = False
        If ObjModuloPorServicioBln.ObjValorPro Then
            lblnPresOk = ObjValorPres_AnoDec.ObjValorPro > 0
        Else
            For Each lobjSer As ClsServicio In ColServiciosAno
                If Not lobjSer.ObjEsAjusteBln.ObjValorPro Then
                    lblnPresOk = If(lobjSer.ObjMiAno.ObjTipoCalculoCuotaByt.ObjValorPro <>
                            EnuTipoBaseCalculo.EnuImportadas, lobjSer.FdecValor > 0,
                            lobjSer.FdecValor = 0)
                    If Not lblnPresOk Then
                        If (lobjSer.ObjMiAno.ObjTipoCalculoCuotaByt.ObjValorPro =
                                EnuTipoBaseCalculo.EnuCoeficientePro AndAlso
                                lobjSer.FdecValor = 0) Then
                            astrMens = "Se debe ingresar los valores de los módulos de contribuciòn!"
                        End If
                        Exit For
                    End If
                End If
            Next
        End If
        Return lblnPresOk
    End Function
    Friend Function FblnEstaAjustada() As Boolean
        Dim lblnEstaAju = False
        If ColServiciosAno.Count > 0 Then
            For Each lobjSer As ClsServicio In ColServiciosAno
                If Not lobjSer.ObjEsAjusteBln.ObjValorPro Then
                    lblnEstaAju = lobjSer.ObjEstaAjustadoBln.ObjValorPro
                End If
                If Not lblnEstaAju Then Exit For
            Next
        End If
        Return lblnEstaAju
    End Function
    Private Sub SRegisCalculo()
        For Each lobjSer As ClsServicio In ColServiciosAno
            lobjSer.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
            lobjSer.ObjEstaGenaradaProgramBln.ObjValorPro = True
            lobjSer.SActualice(True)
        Next
    End Sub
    ''' <summary>
    ''' Indica si las cuotas de administración provisionales pueden ser calculadas
    ''' </summary>
    ''' <returns></returns>
    Friend Function FblnPuedeCalcularCuotasIni() As Boolean
        Return (ObjValorPres_AnoDec.ObjValorPro > 0)
    End Function
    Friend Function FblnDebeCalcularCuotas() As Boolean
        Dim lblnDebeCalcularCuotas = (ObjIdAnoShr.ObjValorPro >=
                GobjParametros.ObjAnoActual.ObjIdAnoShr.ObjValorPro)
        If lblnDebeCalcularCuotas Then
            lblnDebeCalcularCuotas = ObjValorPres_AnoDec.ObjValorPro > 0
        End If
        If lblnDebeCalcularCuotas Then
            For Each lobjServicio As ClsServicio In ColServiciosAno
                If Not lobjServicio.ObjEsAjusteBln.ObjValorPro Then
                    lblnDebeCalcularCuotas = Not lobjServicio.ObjEstaGenaradaProgramBln.ObjValorPro
                    If lblnDebeCalcularCuotas Then Exit For
                End If
            Next
        End If
        Return lblnDebeCalcularCuotas
    End Function
    Friend Function FblnDebeAjustarCuotasAdmin() As Boolean
        Dim lblnDebeAjustar = (ObjIdAnoShr.ObjValorPro =
                GobjParametros.ObjAnoActual.ObjIdAnoShr.ObjValorPro) AndAlso
                Not FblnEstaAjustada() AndAlso FblnEstaGenCuota()
        If lblnDebeAjustar Then
            lblnDebeAjustar = ClsOrionCop.FblnHayPeriFacturados AndAlso
                    Not GobjParametros.ObjAnoActual.FblnFacturacionGenerada
        End If
        Return lblnDebeAjustar
    End Function
#End Region
End Class

#Region "Clases de Propiedad"
Friend Class ClsDiasParaDsctoPPShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "DiasDescuentoPP"
    Private ReadOnly MobjPadre As ClsAno = Nothing

    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Dias para descuento Pronto Pago"
        HStrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HenuTipoValor = EnuTipoValor.EnuShort
        HblnRegistrarLogCambio = True
    End Sub

    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        HblnEsRequerido = MobjPadre.ObjTipoIncentivoByt.ObjValorPro =
                EnuTipoIncentivo.EnuDescuentoPP
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1,
                Short.MaxValue, BlnEsRequerido, EnuTipoValor)
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

Friend Class ClsDiasMultaExtShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "DiasPExtemporaneo"
    Private ReadOnly MobjPadre As ClsAno = Nothing

    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Dias para pago extemporáneo"
        HStrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HenuTipoValor = EnuTipoValor.EnuShort
        HblnRegistrarLogCambio = True
    End Sub

    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        HblnEsRequerido = MobjPadre.ObjTipoIncentivoByt.ObjValorPro =
                EnuTipoIncentivo.EnuPenalización
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1,
                30, BlnEsRequerido, EnuTipoValor)
        If HblnEsValido AndAlso HobjValorNew > 0 Then
            HblnEsValido = MobjPadre.ObjTipoIncentivoByt.ObjValorPro =
                    EnuTipoIncentivo.EnuPenalización
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

Friend Class ClsEstaCerradoAnoBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "EstaCerrado"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "EstaCerradoAño"
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HenuTipoValor = EnuTipoValor.enuBoolean
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class

Friend Class ClsFechaUltEstadoCtaDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaUltimoEstadoCta"
    Private ReadOnly MobjPadre As ClsAno = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Fecha último Estado Cuenta"
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
            Dim ldtmFechaMin As Date = GCDTMFECHANULA
            Dim ldtmFechaMax As Date = Date.Today.AddDays(-1)
            HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax, BlnEsRequerido)
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

Friend Class ClsIdAnoShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdAno"
    Private ReadOnly MobjPadre As ClsAno = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdAño"
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HenuTipoValor = EnuTipoValor.enuShort
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 2
    End Sub
    Public Overrides Sub SValide()
        HstrMens = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCSHRANOMINIMO, GCSHRANOMAXIMO,
                BlnEsRequerido, EnuTipoValor)
        If HblnEsValido Then
            If Not BlnLeyendoOrigen Then
                If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                    If HblnEsValido Then
                        Dim lobjValorLlave() = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
                        If Not MobjPadre.FblnExisteLlave(lobjValorLlave) Then
                            HstrMens = "El  Año ingresado, " & HobjValorNew.ToString & "', no existe!"
                            HblnEsValido = False
                        End If
                    End If
                ElseIf MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                    HblnEsValido = (HobjValorOriginal = HobjValorNew)
                    If Not HblnEsValido Then
                        HstrMens = "No es permitido cambiar la identidad a objeto alguno!"
                    End If
                End If
            End If
        Else
            HstrMens = "El Año ingresado, " & HobjValorNew.ToString & "', no es válido!"
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
            Return String.Empty
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsModuloPorServicioBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "ModuloPorServicio"
    Private ReadOnly MobjPadre As ClsAno = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Módulo por Servicio"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub

    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
        If Not BlnLeyendoOrigen Then
            If HblnEsValido Then
                If HobjValorNew Then
                    For Each lobjServicio As ClsServicio In MobjPadre.ColServiciosAno
                        If Not lobjServicio.ObjEsAjusteBln.ObjValorPro AndAlso
                                lobjServicio.ColModulosServicio.Count > 1 Then
                            HblnEsValido = False
                            HstrMens = "Hay al menos un Servicio Cuota de Administración " &
                                    "con más de un Módulo de Contribución asociado!"
                            SNotifiqueDatInv()
                            Exit For
                        End If
                    Next
                    If HblnEsValido AndAlso HobjValorNew = True Then
                        HblnEsValido = Not ClsOrionCop.FblnHaySectContriAdminMasDeUnaVes
                        If Not HblnEsValido Then
                            HstrMens = "Hay un Sector que contribuye a más de un Módulo de Administración"
                            SNotifiqueDatInv()
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
            MobjPadre.ObjValorPres_AnoDec.SValide()
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

Friend Class ClsIdServicioMultaShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdServicioMulta"
    Private ReadOnly MobjPadre As ClsAno = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Servicio Multa"
        HenuTipoValor = EnuTipoValor.EnuShort
        HStrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub

    Public Overrides Sub SValide()
        HblnEsRequerido = MobjPadre.ObjTipoIncentivoByt.ObjValorPro =
                EnuTipoIncentivo.EnuPenalización
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue,
                BlnEsRequerido, EnuTipoValor.EnuShort)
        If Not BlnLeyendoOrigen Then
            If HblnEsValido AndAlso MobjPadre.EnuEstadoActualizacion <>
                    EnuEstadoObjetoDef.EnuConsultando Then
                If HblnEsValido AndAlso BlnEsRequerido Then
                    Dim LstrKey = "0," & HobjValorNew.ToString()
                    HblnEsValido = GobjParametros.ColServiciosPer.Contains(LstrKey)
                    If Not HblnEsValido Then
                        HstrMens = "El Servicio ingresado para la Multa no existe!"
                        SNotifiqueDatInv()
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando Then
            MobjPadre.ObjValorPres_AnoDec.SValide()
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

Friend Class ClsTipoCalculoCuotaByt
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsAno = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTipoCalculoCuota"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Tipo Calculo Cuota Admin"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        HobjValorNew = EnuTipoBaseCalculo.None
        HobjValorPro = EnuTipoBaseCalculo.None
    End Sub
    Public Overrides Sub SValide()
        HstrMens = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew,
                EnuTipoBaseCalculo.EnuCoeficientePro,
                EnuTipoBaseCalculo.EnuImportadas, BlnEsRequerido)
        If HblnEsValido Then
            Dim lblnActuServiciosEstanCalculados = HobjValorOriginal =
                    EnuTipoBaseCalculo.EnuImportadas
            If HobjValorNew = EnuTipoBaseCalculo.EnuCuotaAnterior Then
                Dim lshrIdAno = MobjPadre.ObjIdAnoShr.ObjValorPro
                HblnEsValido = Not GobjParametros.FblnAnoEsElPrimero(lshrIdAno)
                If Not HblnEsValido Then
                    HstrMens = "No es posible porque el Año anterior no tiene Cuotas Calculadas!"
                End If
            End If
            If HobjValorNew = EnuTipoBaseCalculo.EnuImportadas Then
                If ObjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                    If HobjValorNew <> HobjValorOriginal Then
                        HblnEsValido = GblnImportando
                    End If
                    If Not HblnEsValido Then
                        HstrMens = "Este valor es determinado automáticamente cuando las cuotas " &
                            "se importan"
                        SNotifiqueDatInv()
                    End If
                ElseIf ObjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                    HblnEsValido = False
                    HstrMens = "Este valor es determinado automáticamente cuando las cuotas " &
                            "se importan"
                    SNotifiqueDatInv()
                End If
            End If
            If HobjValorNew <> HobjValorOriginal AndAlso lblnActuServiciosEstanCalculados Then
                For Each lobjServicio As ClsServicio In MobjPadre.ColServiciosAno
                    If lobjServicio.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                        lobjServicio.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                    End If
                    lobjServicio.ObjGeneraProgramBln.ObjValorPro = True
                    lobjServicio.ObjEstaGenaradaProgramBln.ObjValorPro = False
                    lobjServicio.ObjEstaAjustadoBln.ObjValorPro = False
                    lobjServicio.ObjTipoBaseCalculoByt.ObjValorPro =
                            MobjPadre.ObjTipoCalculoCuotaByt.ObjValorPro
                Next
            End If
        Else
            HstrMens = "El Dato ingresado no es válido!"
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
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsTipoIncentivoByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTipoIncentivo"
    Private ReadOnly MobjPadre As ClsAno = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Tipo Incentivo al pago"
        HenuTipoValor = EnuTipoValor.EnuByte
        HStrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew,
                EnuTipoIncentivo.None, EnuTipoIncentivo.EnuPenalización, BlnEsRequerido)
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

Friend Class ClsTipoDsctoPPByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTipoDsctoPP"
    Private ReadOnly MobjPadre As ClsAno = Nothing

    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Tipo Descuento Pronto Pago"
        HenuTipoValor = EnuTipoValor.EnuByte
        HStrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub

    Protected Overrides Sub SVaciePropiedad()
        HobjValorNew = EnuTipoDsctoPP.None
        HobjValorPro = EnuTipoDsctoPP.None
    End Sub

    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew,
                EnuTipoDsctoPP.None, EnuTipoDsctoPP.EnuValorFijo, BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando Then
                If MobjPadre.ObjTipoIncentivoByt.ObjValorPro =
                        EnuTipoIncentivo.EnuDescuentoPP Then
                    HblnEsValido = (HobjValorNew > EnuTipoDsctoPP.None) AndAlso
                            (HobjValorNew <= EnuTipoDsctoPP.EnuValorFijo)
                Else
                    HblnEsValido = (HobjValorNew = EnuTipoDsctoPP.None)
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
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsValorMultaPagoExtDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "ValorMultaPExtemporaneo"
    Private ReadOnly MobjPadre As ClsAno = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Valor Multa Pago Extemporáneo"
        HenuTipoValor = EnuTipoValor.EnuDecimal
        HStrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        HblnEsRequerido = MobjPadre.ObjTipoIncentivoByt.ObjValorPro =
                EnuTipoIncentivo.EnuPenalización
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1,
                    Decimal.MaxValue, BlnEsRequerido, HenuTipoValor)
        If Not HblnEsValido Then
            HstrMens = "El valor de la multa por pago extemporáneo debe ser mayor a cero!"
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
End Class

Friend Class ClsValorPres_AnoDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "ValorPresIng"
    Private ReadOnly MobjPadre As ClsAno = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Valor Presupuesto Año"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        HblnEsValido = True
        If Not BlnLeyendoOrigen Then
            HstrMens = String.Empty
            HblnEsRequerido = MobjPadre.ColServiciosAno.Count > 0
            HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0,
                    Decimal.MaxValue, BlnEsRequerido, HenuTipoValor)
            If HblnEsValido Then
                If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                    If MobjPadre.FblnValorPresExigeCero Then
                        HblnEsValido = (HobjValorNew = 0)
                    Else
                        If MobjPadre.ObjModuloPorServicioBln.ObjValorPro AndAlso
                                BlnEsRequerido AndAlso (GobjParametros.EnuEstadoAplicacion =
                                EnuEstadoAplicacionDef.EnuNormal OrElse
                                GobjParametros.EnuEstadoAplicacion =
                                EnuEstadoAplicacionDef.EnuSinPresupuesto) Then
                            HblnEsValido = (HobjValorNew > 0)
                            If Not HblnEsValido Then
                                HstrMens = "El Presupuesto debe tener un Valor mayor a Cero!"
                            End If
                        End If
                    End If
                End If
            End If
            If Not String.IsNullOrEmpty(HstrMens) Then
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
            Return Format(HobjValorPro, "c")
        End If
    End Function
End Class
#End Region
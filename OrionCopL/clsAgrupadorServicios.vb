Friend Class ClsAgrupadorServicios
#Region "Definiciones"
    Inherits clsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriAgrupadoresServicios"
    '
    Private McolServiciosAgrupados As Collection = Nothing
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia un objeto Panorama.
    ''' </summary>
    ''' <param name="aenuModoInstanciaObj">Indica si se instancia como un objeto navegable o como un Objeto único.</param>
    ''' <remarks>Si se instancia como un objeto navegable, se crea un datatable que contiene las columnas de
    ''' la llave con las llaves de todos los objetos y queda a la espera de que se indique que objeto abrir.
    ''' Si se instancia como un objeto único, queda a la espera de recibir el valor de los campos de la llave 
    ''' para abrir dicho objeto. </remarks>
    Public Sub New(aenuModoInstanciaObj As EnuModoInstanciaObjDef)
        If aenuModoInstanciaObj <> EnuModoInstanciaObjDef.enuDeColeccion Then
            HobjPadre = Nothing
            Dim lstrCamposSelect As String()
            If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuNavegable Then
                HblnEsAnulable = False
                HblnEsSuprimible = False
                HcolFiltros.Add(ClsOrionCop.StrFiltroUbicacion)
                lstrCamposSelect = {ClsIdCarpetaShr.SstrNombreCampoBd, ClsIdCentroUtilShr.SstrNombreCampoBd,
                    ClsIdAgrupadorServiciosShr.SstrNombreCampoBd}
            Else
                HblnEsCreable = False
                HblnEsModificable = False
                HblnEsSuprimible = False
                HblnEsAnulable = False
                HenuTipoObjeto = EnuModoInstanciaObjDef.enuUnico
                lstrCamposSelect = {"*"}
            End If
            HcolTablas.Add(MCSTRNOMBRETABLA)
            HcolCamposSelect.Add(lstrCamposSelect)
        Else
            Throw New ErrorInesperadoPanLException("Con este Constructor no se puede instanciar un Objeto de Colección!")
        End If
    End Sub
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Administrador al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwObjeto">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Friend Sub New(aobjPadre As Object, adrwObjeto As DataRow)
        HobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsAnulable = False
        HblnEsSuprimible = False
        '
        DrwRegistroActual = adrwObjeto
        DtbTablaColeccion = DrwRegistroActual.Table
    End Sub
#End Region
#Region "Propiedades"
#Region "Propiedades identificadoras"
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
            Return EnuIdClasesPanDef.enuAgrServicios
        End Get
    End Property

    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Agrupador De Servicios"
        End Get
    End Property

    Friend Overrides ReadOnly Property HstrNombreObj As String
        Get
            Return Chr(34) & ObjNombreAgrupadorServiciosStr.ObjValorPro & Chr(34)
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjDiaFacturaShr As New ClsDiaFacturaShr(Me)
    Friend ReadOnly Property ObjDiasGraciaShr As New ClsDiasGraciaShr(Me)
    Friend ReadOnly Property ObjDiasVencimientoShr As New ClsDiasVencimientoShr(Me)
    Friend ReadOnly Property ObjFactAPropSinPreAgrBln As New ClsFactAPropSinPreAgrBln(Me)
    Friend ReadOnly Property ObjFactAPropYPreAgrBln As New ClsFactAPropYPreAgrBln(Me)
    Friend ReadOnly Property ObjIdAgrupadorServiciosShr As New ClsIdAgrupadorServiciosShr(Me)
    Friend ReadOnly Property ObjIdCarpetaAgruServiciosShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtilAgruServiciosShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjNombreAgrupadorServiciosStr As New ClsNombreAgrupadorServiciosStr(Me)
    Friend ReadOnly Property ObjPeriodoGraciaFinMesBln As New ClsPeriodoGraciaFinMesBln(Me)
    Friend ReadOnly Property ObjPieFacturaDosStr As New ClsPieFacturaDosStr(Me)
    Friend ReadOnly Property ObjPieFacturaUnoStr As New ClsPieFacturaUnoStr(Me)
    Friend ReadOnly Property ObjVenceFinMesBln As New ClsVenceFinMesBln(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjDiaFacturaShr)
                HcolPropiedades.Add(ObjDiasGraciaShr)
                HcolPropiedades.Add(ObjDiasVencimientoShr)
                HcolPropiedades.Add(ObjFactAPropSinPreAgrBln)
                HcolPropiedades.Add(ObjFactAPropYPreAgrBln)
                HcolPropiedades.Add(ObjIdAgrupadorServiciosShr)
                HcolPropiedades.Add(ObjIdCarpetaAgruServiciosShr)
                HcolPropiedades.Add(ObjIdCentroUtilAgruServiciosShr)
                HcolPropiedades.Add(ObjNombreAgrupadorServiciosStr)
                HcolPropiedades.Add(ObjPeriodoGraciaFinMesBln)
                HcolPropiedades.Add(ObjPieFacturaDosStr)
                HcolPropiedades.Add(ObjPieFacturaUnoStr)
                HcolPropiedades.Add(ObjVenceFinMesBln)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    Friend ReadOnly Property DtmFechaFacturacionPeriodoActual As Date
        Get
            Dim ldtmFechaFact = GCDTMFECHANULA
            If ObjDiaFacturaShr.BlnEsValido Then
                Dim lobjAnoActual = GobjCentroUtilOriCop.ObjAnoActual
                If Not IsNothing(lobjAnoActual) Then
                    Dim lobjPeriodoActual As ClsPeriodo = lobjAnoActual.ObjPeriodoActual
                    If Not IsNothing(lobjPeriodoActual) Then
                        If GobjCentroUtilOriCop.ObjExigeFechaHoyDocsBln.ObjValorPro Then
                            ldtmFechaFact = Date.Today
                        Else
                            Dim lshrAno As Short = lobjAnoActual.ObjIdAnoShr.ObjValorPro
                            Dim lshrMes As Short = lobjPeriodoActual.ObjIdPeriodoShr.ObjValorPro
                            ldtmFechaFact = DateSerial(lshrAno, lshrMes, ObjDiaFacturaShr.ObjValorPro)
                        End If
                    End If
                End If
            End If
            Return ldtmFechaFact
        End Get
    End Property
    Friend ReadOnly Property DtmFechaVencePeriActual As Date
        Get
            Dim ldtmFechaVence = GCDTMFECHANULA
            If ObjDiasVencimientoShr.BlnEsValido Then
                Dim lobjAnoActual = GobjCentroUtilOriCop.ObjAnoActual
                If Not IsNothing(lobjAnoActual) Then
                    Dim lobjPeriodoActual As ClsPeriodo = lobjAnoActual.ObjPeriodoActual
                    If Not IsNothing(lobjPeriodoActual) Then
                        If ObjVenceFinMesBln.ObjValorPro Then
                            ldtmFechaVence = lobjPeriodoActual.DtmFechaFinPeriodo
                        Else
                            ldtmFechaVence = DtmFechaFacturacionPeriodoActual.AddDays(ObjDiasVencimientoShr.ObjValorPro)
                        End If
                    End If
                End If
            End If
            Return ldtmFechaVence
        End Get
    End Property
    Friend ReadOnly Property DtmFechaGraciaPeriActual As Date
        Get
            Dim ldtmFechaGracia = GCDTMFECHANULA
            If ObjDiasGraciaShr.BlnEsValido Then
                Dim lobjAnoActual = GobjCentroUtilOriCop.ObjAnoActual
                If Not IsNothing(lobjAnoActual) Then
                    Dim lobjPeriodoActual As ClsPeriodo = lobjAnoActual.ObjPeriodoActual
                    If Not IsNothing(lobjPeriodoActual) Then
                        If ObjPeriodoGraciaFinMesBln.ObjValorPro Then
                            Dim lentMesFV = DtmFechaVencePeriActual.Month
                            Dim lentAnoFV = DtmFechaVencePeriActual.Year
                            Dim lobjAno As ClsAno = GobjCentroUtilOriCop.ColAnos(lentAnoFV.ToString)
                            Dim lobjPeriodoGracia As ClsPeriodo = lobjAno.ColPeriodos(Format(lentMesFV, "0#"))
                            ldtmFechaGracia = lobjPeriodoGracia.DtmFechaFinPeriodo
                        Else
                            ldtmFechaGracia = DtmFechaVencePeriActual.AddDays(ObjDiasGraciaShr.ObjValorPro)
                        End If
                    End If
                End If
            End If
            Return ldtmFechaGracia
        End Get
    End Property
    Friend ReadOnly Property ColServiciosAgrupados As Collection
        Get
            'If IsNothing(McolServiciosAgrupados) Then
            '    Dim lstrKey As String
            '    McolServiciosAgrupados = New Collection
            '    For Each lobjServicio As ClsServicio In GobjCentroUtilOriCop.ObjAnoActual.ColServiciosAno
            '        If lobjServicio.ObjIdAgruEsteServicioShr.ObjValorPro =
            '                ObjIdAgrupadorServiciosShr.ObjValorPro Then
            '            lstrKey = lobjServicio.ObjIdAno_ServicioShr.ToString & "," &
            '                    lobjServicio.ObjIdServicioShr.ToString
            '            McolServiciosAgrupados.Add(lobjServicio, lstrKey)
            '        End If
            '    Next
            '    For Each lobjServicio As ClsServicio In GobjCentroUtilOriCop.ColServiciosPer
            '        If lobjServicio.ObjIdAgruEsteServicioShr.ObjValorPro =
            '                ObjIdAgrupadorServiciosShr.ObjValorPro Then
            '            lstrKey = lobjServicio.ObjIdAno_ServicioShr.ToString & "," &
            '                    lobjServicio.ObjIdServicioShr.ToString
            '            McolServiciosAgrupados.Add(lobjServicio, lstrKey)
            '        End If
            '    Next
            'End If
            Return McolServiciosAgrupados
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SInicialiceObj()
        ObjIdCarpetaAgruServiciosShr.ObjValorPro = GshrIdCarpeta
        ObjIdCentroUtilAgruServiciosShr.ObjValorPro = GshrIdCentroUtil
        ObjDiaFacturaShr.ObjValorPro = 1
        ObjDiasVencimientoShr.ObjValorPro = 9
        ObjFactAPropYPreAgrBln.ObjValorPro = False
        ObjFactAPropSinPreAgrBln.ObjValorPro = False
        ObjIdAgrupadorServiciosShr.ObjValorPro = 0
        ObjNombreAgrupadorServiciosStr.ObjValorPro = String.Empty
        ObjPeriodoGraciaFinMesBln.ObjValorPro = False
        ObjPieFacturaDosStr.ObjValorPro = String.Empty
        ObjPieFacturaUnoStr.ObjValorPro = String.Empty
        ObjVenceFinMesBln.ObjValorPro = False
        McolServiciosAgrupados = Nothing
    End Sub
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        GobjPanDat.SControleProcesoObj(True)
        Try
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                SNumereObj()
            End If
            If EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                Dim lblnAdicioneCol = EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando
                If ObjIdAgrupadorServiciosShr.ObjValorPro <> 1 Then
                    SGeneralicePiePagina()
                End If
                MyBase.SActualice(ablnExigeRequeridos)
                'GobjCentroUtilOriCop.SRefesqueAgruServicio()
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
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjIdAgrupadorServiciosShr.ToString
        End Get
    End Property
    Protected Overrides Function FblnSuprimio() As Boolean
        Dim lblnSuprimio = FblnEsSuprimible()
        If lblnSuprimio Then
            lblnSuprimio = MyBase.FblnSuprimio()
            If lblnSuprimio Then
                If BlnEsNavegable Then
                    GobjPanorama.SRegistreAccionLogApp(HstrNombreClase,
                            "Suprimir Agrupador de Ser. " &
                            ObjIdAgrupadorServiciosShr.ToString & "-" &
                            ObjNombreAgrupadorServiciosStr.ObjValorPro)
                End If
            End If
        End If
        Return lblnSuprimio
    End Function
    Friend Overrides Function FblnEsSuprimible() As Boolean
        Dim lblnEsSuprimible = MyBase.FblnPermitidoSuprimir()
        If lblnEsSuprimible Then
            Dim lstrCond = " = " & ObjIdAgrupadorServiciosShr.ToString & " AND " &
                    ClsOrionCop.StrFiltroUbicacion
            lblnEsSuprimible = ClsPanorama.FblnEsEliminableReg({SstrNombreTabla},
                    ClsIdAgrupadorServiciosShr.SstrNombreCampoBd, lstrCond, True, False)
        End If
        Return lblnEsSuprimible
    End Function
#End Region
#Region "Procedimientos del objeto"
    Private Sub SNumereObj()
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            Dim lshrIdAgrSer As Short
            lshrIdAgrSer = ClsPanorama.FobjUltimaIdNumericaObjeto(MCSTRNOMBRETABLA,
                    ClsIdAgrupadorServiciosShr.SstrNombreCampoBd, ObjIdAgrupadorServiciosShr.EnuTipoValor,
                    ClsOrionCop.StrFiltroUbicacion) + 1
            ObjIdAgrupadorServiciosShr.ObjValorPro = lshrIdAgrSer
            Dim lstrPieFacUno = GobjCentroUtilOriCop.ObjPieFacturaUnoStr.ObjValorPro
            If Not String.IsNullOrEmpty(lstrPieFacUno) Then
                ObjPieFacturaUnoStr.ObjValorPro = lstrPieFacUno
            End If
        End If
    End Sub
    ''' <summary>
    ''' Indica si hay items del programa de facturacion que correspondan a un servicio 
    ''' de este agrupador para ser facturados en el periodo actual.
    ''' </summary>
    Friend Function FblnHayItemsPorFacturar() As Boolean
        Dim lblnHay = False
        For Each lobjServicio As ClsServicio In ColServiciosAgrupados
            lblnHay = FblnHayItemsSerPorFacturar(lobjServicio)
            If lblnHay Then Exit For
        Next
        Return lblnHay
    End Function
    Private Function FblnHayItemsSerPorFacturar(aobjServicio As ClsServicio) As Boolean
        Dim lblnPorFact = False
        Dim lentPerPorFac As Integer, lentTotCantPer As Integer, lentCantPeriFacturados As Integer, lstrPeriodoIni As String
        Dim lstrPeriodoFacturado As String, lstrPeriodoActual As String
        Dim lshridAno As Short = aobjServicio.ObjIdAno_ServicioShr.ObjValorPro
        Dim lshrIdServicio As Short = aobjServicio.ObjIdServicioShr.ObjValorPro
        Dim lstrFiltro = ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " = " & lshridAno & " AND " &
                ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd & " = " & lshrIdServicio
        Dim ldrwItems As DataRow() = ClsOrionCop.FdtbItemsProgFactConSaldo(False).Select(lstrFiltro)
        For Each ldrwItem As DataRow In ldrwItems
            lentPerPorFac = ClsPanorama.FobjValorCampo(ldrwItem("CantPerPorFact"),
                  EnuTipoValorDef.enuInteger)
            lentTotCantPer = ClsPanorama.FobjValorCampo(ldrwItem(
                  ClsCantidadPeriodosShr.SstrNombreCampoBd), EnuTipoValorDef.enuInteger)
            lstrPeriodoIni = ClsPanorama.FobjValorCampo(ldrwItem(
                  ClsPeriodoIni_ItemProgStr.SstrNombreCampoBd), EnuTipoValorDef.enuString)
            lentCantPeriFacturados = lentTotCantPer - lentPerPorFac
            lstrPeriodoFacturado = ClsOrionCop.FstrPeriodoFinal(lstrPeriodoIni, lentCantPeriFacturados - 1)
            lstrPeriodoActual = GobjCentroUtilOriCop.ObjAnoActual.StrIdPeriodoActual
            If lstrPeriodoFacturado < lstrPeriodoActual Then
                Dim lstrPeriodoHoy = ClsOrionCop.FstrPeriodoDeFecha(Date.Today)
                If lstrPeriodoHoy > lstrPeriodoActual Then
                    lblnPorFact = True
                Else
                    lblnPorFact = (Day(DtmFechaFacturacionPeriodoActual) <= Day(Date.Today))
                End If
            End If
            If lblnPorFact Then Exit For
        Next
        Return lblnPorFact
    End Function
    Private Sub SGeneralicePiePagina()
        'Dim lcolAgrSer = GobjCentroUtilOriCop.ColAgrupadoresServicios
        'Dim lstrPieFraUno_Admin = String.Empty, lstrPieFraDos_Admin = String.Empty
        'If lcolAgrSer.Count > 0 Then
        '    Dim lobjAgrSer As ClsAgrupadorServicios = lcolAgrSer("1")
        '    lstrPieFraUno_Admin = lobjAgrSer.ObjPieFacturaUnoStr.ToString
        '    lstrPieFraDos_Admin = lobjAgrSer.ObjPieFacturaDosStr.ToString
        'End If
        'ObjPieFacturaUnoStr.ObjValorPro = lstrPieFraUno_Admin
        'If String.IsNullOrEmpty(ObjPieFacturaDosStr.ToString()) Then
        '    ObjPieFacturaDosStr.ObjValorPro = lstrPieFraDos_Admin
        'End If
    End Sub
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsDiaFacturaShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "DiaCorte"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "DiaDeCorte"
        HenuTipoValor = EnuTipoValorDef.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1,
                31, BlnEsRequerido, HenuTipoValor)
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
Friend Class ClsDiasGraciaShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "DiasGracia"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "DiasDeGracia"
        HenuTipoValor = EnuTipoValorDef.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsAgrupadorServicios = ObjPadre
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0,
                Short.MaxValue, BlnEsRequerido, HenuTipoValor)
        If lblnEsValido Then
            If lobjPadre.ObjPeriodoGraciaFinMesBln.ObjValorPro Then
                lblnEsValido = (HobjValorNew = 0)
            End If
        End If
        HblnEsValido = lblnEsValido
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
Friend Class ClsDiasVencimientoShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "DiasVencimiento"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "DiasVencimiento"
        HenuTipoValor = EnuTipoValorDef.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsAgrupadorServicios = ObjPadre
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0,
                Short.MaxValue, BlnEsRequerido, HenuTipoValor)
        If HblnEsValido Then
            If lobjPadre.ObjVenceFinMesBln.ObjValorPro Then
                HblnEsValido = (HobjValorNew = 0)
            Else
                HblnEsValido = (HobjValorNew > 0)
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
Friend Class ClsFactAPropYPreAgrBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FacturarSoloPropietario"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "FacturarAPropietario"
        HenuTipoValor = EnuTipoValorDef.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
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
        Return ClsPanorama.FstrBuleanoToString(HobjValorPro)
    End Function
End Class
Friend Class ClsFactAPropSinPreAgrBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FacturarSoloPropSinPreAgr"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Facturar A Propietario Sin Predio Agr."
        HenuTipoValor = EnuTipoValorDef.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
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
        Return ClsPanorama.FstrBuleanoToString(HobjValorPro)
    End Function
End Class
Friend Class ClsIdAgrupadorServiciosShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdAgrupadorServicios"
    Private ReadOnly MobjPadre As ClsAgrupadorServicios = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdAgrupadorServicios"
        HenuTipoValor = EnuTipoValorDef.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HstrOrdenIndice = "ASC"
        HbytPosicionLlave = 2
    End Sub
    Public Overrides Sub SValide()
        Dim lobjValorIng As Object = HobjValorNew
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue,
                BlnEsRequerido, EnuTipoValor)
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
            If HblnEsValido Then
                If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                    If Not BlnLeyendoOrigen Then
                        Dim lobjLlavePrincipal() = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
                        If Not MobjPadre.FblnExisteLlave(lobjLlavePrincipal) Then
                            HblnEsValido = False
                            HstrMens = "La Id. del Agrupador de Servicios ingresada, '" &
                                    lobjValorIng.ToString & "', no existe!"
                            SLevanteEveNot("", 0, EnuSeveridadNot.EnuDatoInvalido)
                        End If
                    End If
                ElseIf MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                    HblnEsValido = (HobjValorOriginal = HobjValorNew)
                    If Not HblnEsValido Then
                        HstrMens = "No es permitido cambiar la identidad a objeto alguno!"
                        SLevanteEveNot("", 0, EnuSeveridadNot.EnuDatoInvalido)
                    End If
                End If
            ElseIf Not IsNothing(lobjValorIng) Then
                HstrMens = "La Id. del Agrupador de Servicios ingresada, '" &
                        lobjValorIng.ToString & "', no es válida!"
                SLevanteEveNot("", 0, EnuSeveridadNot.EnuDatoInvalido)
            End If
        Else
            HblnEsValido = True
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
Friend Class ClsNombreAgrupadorServiciosStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "NombreAgrupadorServicios"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "NombreAgrupadorServicios"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValorDef.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 3, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            HobjValorNew = HobjValorNew.ToString.ToUpper
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
Friend Class ClsPeriodoGraciaFinMesBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PeriodoGraciaFinMes"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "GraciaFinMes"
        HenuTipoValor = EnuTipoValorDef.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        Dim lobjPadre As ClsAgrupadorServicios = ObjPadre
        If HblnEsValido Then
            If CType(HobjValorPro, Boolean) Then
                lobjPadre.ObjDiasGraciaShr.ObjValorPro = 0
            Else
                lobjPadre.ObjDiasGraciaShr.SValide()
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
Friend Class ClsPieFacturaDosStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PieFacturaDos"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "PieFacturaDos"
        HshrLongitud = 230
        HenuTipoValor = EnuTipoValorDef.enuString
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
Friend Class ClsPieFacturaUnoStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PieFacturaUno"
    Private ReadOnly MobjPadre As ClsAgrupadorServicios = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "PieFacturaUno"
        HshrLongitud = 230
        HenuTipoValor = EnuTipoValorDef.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 3, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando AndAlso
                    GobjCentroUtilOriCop.ObjPieFacturaUnoStr.ToString <> "" Then
                If Not String.IsNullOrEmpty(GobjCentroUtilOriCop.FstrPieFacturaRes) Then
                    HblnEsValido = (GobjCentroUtilOriCop.FstrPieFacturaRes = HobjValorNew)
                    If Not HblnEsValido Then
                        HstrMens = "El Pie de Factura uno no puede ser " &
                                "modificado porque está construido a partir " &
                                "Resolución DIAN!"
                        SLevanteEveNot("", 0, EnuSeveridadNot.EnuDatoInvalido)
                    End If
                Else
                    If MobjPadre.ObjIdAgrupadorServiciosShr.ObjValorPro <> 1 Then
                        'HblnEsValido = (GobjCentroUtilOriCop.FstrPieFacturaUno = HobjValorNew)
                        If Not HblnEsValido Then
                            HstrMens = "El Pie de Factura solo puede ser " &
                                    "modificado en el Agrupador ADMINISTRACION!"
                            SLevanteEveNot("", 0, EnuSeveridadNot.EnuDatoInvalido)
                        End If
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
            Return HobjValorPro.ToString().Trim()
        End If
    End Function
End Class
Friend Class ClsVenceFinMesBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "VencimientoFinMes"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "VenceFinMes"
        HenuTipoValor = EnuTipoValorDef.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        Dim lobjPadre As ClsAgrupadorServicios = ObjPadre
        If CType(HobjValorPro, Boolean) Then
            lobjPadre.ObjDiasVencimientoShr.ObjValorPro = 0
        End If
        lobjPadre.ObjDiasVencimientoShr.SValide()
        lobjPadre.ObjPeriodoGraciaFinMesBln.SValide()
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
#End Region
Friend Class ClsServicio
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriServicios"
    ' Variables de modulo
    Private MdtbHistServicio As DataTable = Nothing
    Private McolHistServicio As Collection = Nothing
    Private McolModulosServicio As Collection = Nothing
    Private MobjMiAno As ClsAno = Nothing
    Private MdtmFechaVence As Date = GCDTMFECHANULA
    Private MdtmFechaGracia As Date = GCDTMFECHANULA
    Friend BlnSincronizando As Boolean = False
    Friend BlnCalculandoDesdeAno As Boolean = False
    Friend BlnCreandoAno As Boolean = False
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
    Friend Sub New(aobjAno As ClsAno, aenuModoInstanciaObj As EnuModoInstanciaObjDef)
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuDeColeccion Then
            Throw New ErrorInesperadoPanLException("Con este Constructor no se puede instanciar un Objeto de Colección!")
        End If
        Dim lstrCamposSelect As String()
        HobjPadre = Nothing
        MobjMiAno = aobjAno
        HblnEsSuprimible = False
        HblnEsAnulable = False
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuNavegable Then
            lstrCamposSelect = {ObjIdCarpeta_ServicioShr.StrNombreCampoBD,
                    ObjIdCentroUtil_ServicioShr.StrNombreCampoBD,
                    ObjIdAno_ServicioShr.StrNombreCampoBD,
                    ObjIdServicioShr.StrNombreCampoBD}
            Dim lshrIdAno = 0S
            If Not IsNothing(MobjMiAno) Then
                lshrIdAno = MobjMiAno.ObjIdAnoShr.ObjValorPro
            End If
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsIdAno_ServicioShr.SstrNombreCampoBd & " = " & lshrIdAno
            HcolFiltros.Add(lstrFiltro)
        Else
            HenuTipoObjeto = EnuModoInstanciaObjDef.enuUnico
            lstrCamposSelect = {"*"}
        End If
        HcolTablas.Add(MCSTRNOMBRETABLA)
        HcolCamposSelect.Add(lstrCamposSelect)
    End Sub
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Administrador al cual pertenece el objeto que se está instanciando</param>
    ''' <param name="adrwObjeto">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Friend Sub New(aobjPadre As ClsCBObjetoPan, adrwObjeto As DataRow)
        HobjPadre = aobjPadre
        If TypeOf aobjPadre Is ClsAno Then
            MobjMiAno = aobjPadre
        End If
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
    Friend Shared ReadOnly Property SstrNombreTabla As String
        Get
            Return MCSTRNOMBRETABLA
        End Get
    End Property
    Protected Overrides ReadOnly Property HenuIdClase As EnuIdClasesPanDef
        Get
            Return EnuIdClasesPanDef.EnuServicio
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Servicio"
        End Get
    End Property
    Friend Overrides ReadOnly Property HstrNombreObj As String
        Get
            Return Chr(34) & ObjNombreServicioStr.ObjValorPro & Chr(34)
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjBaseMinimaReteFuenteDec As New ClsBaseMinimaReteFuenteDec(Me)
    Friend ReadOnly Property ObjBaseMinimaReteIcaDec As New ClsBaseMinimaReteIcaDec(Me)
    Friend ReadOnly Property ObjCantPeriodos_ServicioShr As New ClsCantiPeriodos_ServicioShr(Me)
    Friend ReadOnly Property ObjCodigoCuentaCrStr As New ClsCodigoCuentaCrStr(Me)
    Friend ReadOnly Property ObjCodigoCuentaDbStr As New ClsCodigoCuentaDbStr(Me)
    Friend ReadOnly Property ObjCodigoCuentaDevStr As New ClsCodigoCuentaDevStr(Me)
    Friend ReadOnly Property ObjCodigoCuentaIvaStr As New ClsCodigoCuentaIvaStr(Me)
    Friend ReadOnly Property ObjCodigoCuentaMoraStr As New ClsCodigoCuentaMoraStr(Me)
    Friend ReadOnly Property ObjConceptoServicioStr As New ClsConceptoServicioStr(Me)
    ''' <summary>
    ''' Indica si el servicio es un ajuste a un servicio anual, normalmente a la cuota de administración.
    ''' </summary>
    ''' <remarks></remarks>
    Friend ReadOnly Property ObjEsAjusteBln As New ClsEsAjusteBln(Me)
    Friend ReadOnly Property ObjEsExcluidoIvaBln As New ClsEsExcluidoIvaBln(Me)
    ''' <summary>
    ''' Indica si la facturación de este servicio puede ser programada para efectuarse periodicamente.
    ''' </summary>
    ''' <remarks></remarks>
    Friend ReadOnly Property ObjEsFactProgramableBln As New ClsEsFactProgramableBln(Me)
    Friend ReadOnly Property ObjEsServicioIdBln As New ClsEsServicioIdBln(Me)
    Friend ReadOnly Property ObjEstaActivoServicioBln As New ClsEstaActivoServicioBln(Me)
    ''' <summary>
    ''' Indica si un servicio anual que normalmente debe ser ajustado como las Cuotas de Administración,
    '''  ya está ajustado o no.
    ''' </summary>
    ''' <remarks></remarks>
    Friend ReadOnly Property ObjEstaAjustadoBln As New ClsEstaAjustadoBln(Me)
    ''' <summary>
    ''' Indica si ya está generada o no la programación de la facturación.
    ''' </summary>
    ''' <remarks></remarks>
    Friend ReadOnly Property ObjEstaGenaradaProgramBln As New ClsEstaGenaradaProgramBln(Me)
    ''' <summary>
    ''' Indica si la programción para la facturación del servicio será llevada a cabo por Orión,
    '''  asi como el calculo del valor a cobrar por el servicio en cada periodo.
    ''' </summary>
    ''' <remarks></remarks>
    Friend ReadOnly Property ObjGeneraProgramBln As New ClsGeneraProgramBln(Me)
    ''' <summary>
    ''' Indica el año del servicio. Si el valor es cero indica que el servicio es permanente.
    ''' </summary>
    ''' <remarks></remarks>
    Friend ReadOnly Property ObjGraciaFinMesBln As New ClsGraciaFinMes_SerBln(Me)
    Friend ReadOnly Property ObjIdAno_ServicioShr As New ClsIdAno_ServicioShr(Me)
    ''' <summary>
    ''' Indica la Carpeta (Empresa) a la cual pertenece el servicio.
    ''' </summary>
    ''' <remarks></remarks>
    Friend ReadOnly Property ObjIdCarpeta_ServicioShr As New ClsIdCarpetaShr(Me)
    ''' <summary>
    ''' Indica la Copropiedad al cual pertenece el servicio.
    ''' </summary>
    ''' <remarks></remarks>
    Friend ReadOnly Property ObjIdCentroUtil_ServicioShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjModoCausaInteresesByt As New ClsModoCausaInteresesByt(Me)
    ''' <summary>
    ''' Si el servicio es un ajuste, este propiedad identifica del servicio ajustado.
    ''' </summary>
    ''' <remarks></remarks>
    Friend ReadOnly Property ObjIdServicioAjustadoShr As New ClsIdServicioAjustadoShr(Me)
    ''' <summary>
    ''' Identifica el servicio. 
    ''' </summary>
    ''' <remarks></remarks>
    Friend ReadOnly Property ObjIdServicioShr As New ClsIdServicioShr(Me)
    Friend ReadOnly Property ObjIdTerceroCtaCrDbl As New ClsIdTerceroCtaCrDbl(Me)
    ''' <summary>
    ''' Indica si el calculo del valor del servicio se hace con base en el área del predio (:Por Area),
    ''' o por la cantidad de predios (:Por Unidad)
    ''' </summary>
    ''' <remarks></remarks>
    Friend ReadOnly Property ObjTipoBaseCalculoByt As New ClsTipoBaseCalculoByt(Me)
    ''' <summary>
    ''' Determina si el servicio es anual o permanente.
    ''' </summary>
    ''' <remarks></remarks>
    Friend ReadOnly Property ObjIdTipoServicioByt As New ClsIdTipoServicioByt(Me)
    Friend ReadOnly Property ObjIdTipoTerCtaCrSerByt As New ClsIdTipoTerCtaCrSerByt(Me)
    ''' <summary>
    ''' Nombre del servicio. Este nombre será el que aparecerá en el detalle de la factura como el 
    ''' servicio cobrado.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend ReadOnly Property ObjNombreServicioStr As New ClsNombreServicioStr(Me)
    Friend ReadOnly Property ObjPeriodoInicioStr As New ClsPeriodoInicioStr(Me)
    Friend ReadOnly Property ObjTarifaIvaDbl As New ClsTarifaIvaDbl(Me)
    Friend ReadOnly Property ObjTarifaRetFteDbl As New ClsTarifaRetFteDbl(Me)
    Friend ReadOnly Property ObjTarifaRetIcaDbl As New ClsTarifaRetIcaDbl(Me)
    Friend ReadOnly Property ObjDiaFacturaShr As New ClsDiaFactura_SerShr(Me)
    Friend ReadOnly Property ObjDiasGraciaShr As New ClsDiasGracia_SerShr(Me)
    Friend ReadOnly Property ObjDiasVencimientoShr As New ClsDiasVencimiento_SerShr(Me)
    Friend ReadOnly Property ObjFactAPropYPreAgrBln As New ClsFactAPropYPreAgrBln(Me)
    Friend ReadOnly Property ObjVenceFinMesBln As New ClsVenceFinMes_SerBln(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                With HcolPropiedades
                    .Add(ObjIdCarpeta_ServicioShr)
                    .Add(ObjIdCentroUtil_ServicioShr)
                    .Add(ObjIdAno_ServicioShr)
                    .Add(ObjIdServicioShr)
                    .Add(ObjIdServicioAjustadoShr)
                    .Add(ObjDiaFacturaShr)
                    .Add(ObjDiasGraciaShr)
                    .Add(ObjDiasVencimientoShr)
                    .Add(ObjIdTipoServicioByt)
                    .Add(ObjFactAPropYPreAgrBln)
                    .Add(ObjGraciaFinMesBln)
                    .Add(ObjVenceFinMesBln)
                    .Add(ObjBaseMinimaReteFuenteDec)
                    .Add(ObjBaseMinimaReteIcaDec)
                    .Add(ObjCodigoCuentaCrStr)
                    .Add(ObjCodigoCuentaDbStr)
                    .Add(ObjCodigoCuentaDevStr)
                    .Add(ObjCodigoCuentaIvaStr)
                    .Add(ObjCodigoCuentaMoraStr)
                    .Add(ObjConceptoServicioStr)
                    .Add(ObjEsAjusteBln)
                    .Add(ObjEsExcluidoIvaBln)
                    .Add(ObjEsFactProgramableBln)
                    .Add(ObjGeneraProgramBln)
                    .Add(ObjEstaActivoServicioBln)
                    .Add(ObjEstaGenaradaProgramBln)
                    .Add(ObjEsServicioIdBln)
                    .Add(ObjModoCausaInteresesByt)
                    .Add(ObjIdTerceroCtaCrDbl)
                    .Add(ObjTipoBaseCalculoByt)
                    .Add(ObjIdTipoTerCtaCrSerByt)
                    .Add(ObjNombreServicioStr)
                    .Add(ObjPeriodoInicioStr)
                    .Add(ObjTarifaIvaDbl)
                    .Add(ObjTarifaRetFteDbl)
                    .Add(ObjTarifaRetIcaDbl)
                    .Add(ObjCantPeriodos_ServicioShr)
                    .Add(ObjEstaAjustadoBln)
                End With
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    Friend ReadOnly Property BlnFacturaAPropietario As Boolean
        Get
            Return ObjFactAPropYPreAgrBln.ObjValorPro
        End Get
    End Property
    Friend ReadOnly Property DecValor As Decimal
        Get
            Dim ldecValor As Decimal = 0
            For Each lobjModuloServicio As ClsModuloServicio In ColModulosServicio
                ldecValor += lobjModuloServicio.ObjValorPres_ModuloServicioDec.ObjValorPro
            Next
            Return ldecValor
        End Get
    End Property
    Friend ReadOnly Property DecValorAjuste As Decimal
        Get
            Dim ldecAjuste = 0D
            If ObjEsAjusteBln.ObjValorPro Then
                ldecAjuste = ClsOrionCop.FdecValorAjuste(Me)
            End If
            Return ldecAjuste
        End Get
    End Property
    Friend ReadOnly Property DtmFechaUltimaFactura As Date
        Get
            Dim lentDiaFactura As Integer = ObjDiaFacturaShr.ObjValorPro
            Dim lstrPerFin = StrPeriodoFinal
            Dim lentIdAnoFin = CType(Left(lstrPerFin, 4), Integer)
            Dim lentPerFin = CType(Right(lstrPerFin, 2), Integer)
            Dim ldtmfechaUltiFact As Date = DateSerial(lentIdAnoFin, lentPerFin, lentDiaFactura)
            Return ldtmfechaUltiFact
        End Get
    End Property
    Friend ReadOnly Property StrPeriodoFinal As String
        Get
            Dim lstrPerFin = ClsOrionCop.FstrPeriodoFinal(ObjPeriodoInicioStr.ObjValorPro,
                    ObjCantPeriodos_ServicioShr.ObjValorPro - 1)
            Return lstrPerFin
        End Get
    End Property
    Friend ReadOnly Property ObjMiAno As ClsAno
        Get
            If BlnEsCuotaAdministracion Then
                If IsNothing(MobjMiAno) Then
                    MobjMiAno = GobjParametros.ColAnos(ObjIdAno_ServicioShr.ToString)
                End If
            Else
                MobjMiAno = Nothing
            End If
            Return MobjMiAno
        End Get
    End Property
    Friend ReadOnly Property BlnEsCuotaAdministracion As Boolean
        Get
            Dim lblnEs As Boolean
            If ObjIdTipoServicioByt.ObjValorPro IsNot Nothing Then
                lblnEs = If(ObjIdTipoServicioByt.ObjValorPro = EnuTipoServicio.EnuAnual, True, False)
            Else
                lblnEs = MobjMiAno IsNot Nothing
            End If
            Return lblnEs
        End Get
    End Property
    Friend ReadOnly Property DtmFechaFacturacionPeriodoActual As Date
        Get
            Dim ldtmFechaFact = GCDTMFECHANULA
            If ObjDiaFacturaShr.BlnEsValido Then
                Dim lobjAnoActual = GobjParametros.ObjAnoActual
                If Not IsNothing(lobjAnoActual) Then
                    Dim lobjPeriodoActual As ClsPeriodo = lobjAnoActual.ObjPeriodoActual
                    If Not IsNothing(lobjPeriodoActual) Then
                        If GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro Then
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
    Friend ReadOnly Property BlnEsImportado As Boolean
        Get
            Return ObjTipoBaseCalculoByt.ObjValorPro = EnuTipoBaseCalculo.EnuImportadas
        End Get
    End Property
    Friend Property DtmFechaVencePeriActual As Date
        Set(value As Date)
            MdtmFechaVence = value
        End Set
        Get
            If MdtmFechaVence = GCDTMFECHANULA Then
                If ObjDiasVencimientoShr.BlnEsValido Then
                    Dim lobjAnoActual = GobjParametros.ObjAnoActual
                    If Not IsNothing(lobjAnoActual) Then
                        Dim lobjPeriodoActual As ClsPeriodo = lobjAnoActual.ObjPeriodoActual
                        If Not IsNothing(lobjPeriodoActual) Then
                            If ObjVenceFinMesBln.ObjValorPro Then
                                MdtmFechaVence = lobjPeriodoActual.DtmFechaFinPeriodo
                            Else
                                MdtmFechaVence = DtmFechaFacturacionPeriodoActual.AddDays(ObjDiasVencimientoShr.ObjValorPro)
                            End If
                        End If
                    End If
                End If
            End If
            Return MdtmFechaVence
        End Get
    End Property
    Friend Property DtmFechaGraciaPeriActual As Date
        Set(value As Date)
            MdtmFechaGracia = value
        End Set
        Get
            If MdtmFechaGracia = GCDTMFECHANULA Then
                If ObjDiasGraciaShr.BlnEsValido Then
                    Dim lobjAnoActual = GobjParametros.ObjAnoActual
                    If Not IsNothing(lobjAnoActual) Then
                        Dim lobjPeriodoActual As ClsPeriodo = lobjAnoActual.ObjPeriodoActual
                        If Not IsNothing(lobjPeriodoActual) Then
                            If ObjGraciaFinMesBln.ObjValorPro Then
                                Dim lentAnoFV = DtmFechaVencePeriActual.Year
                                Dim lobjAno As ClsAno
                                Dim lobjPeriodoGracia As ClsPeriodo
                                If lentAnoFV > GobjParametros.ObjAnoActual.ObjIdAnoShr.ObjValorPro Then
                                    Dim lentDiasMes = Date.DaysInMonth(DtmFechaVencePeriActual.Year,
                                            DtmFechaVencePeriActual.Month)
                                    MdtmFechaGracia = DateSerial(DtmFechaVencePeriActual.Year,
                                            DtmFechaVencePeriActual.Month, lentDiasMes)
                                Else
                                    Dim lentMesFV = DtmFechaVencePeriActual.Month
                                    lobjAno = GobjParametros.ColAnos(lentAnoFV.ToString)
                                    lobjPeriodoGracia = lobjAno.ColPeriodos(Format(lentMesFV, "0#"))
                                    MdtmFechaGracia = lobjPeriodoGracia.DtmFechaFinPeriodo
                                End If
                            ElseIf ObjDiasGraciaShr.ObjValorPro = 0 Then
                                MdtmFechaGracia = MdtmFechaVence
                            Else
                                MdtmFechaGracia =
                                        DtmFechaVencePeriActual.AddDays(ObjDiasGraciaShr.ObjValorPro)
                            End If
                        End If
                    End If
                End If
            End If
            Return MdtmFechaGracia
        End Get
    End Property
    ''' <summary>
    ''' Fecha de gracia cuando la factura es manual y no auromática
    ''' </summary>
    ''' <param name="adtmFechaVence"></param>
    ''' <returns></returns>
    Friend Function FdtmFechaGracias(adtmFechaVence As Date) As Date
        Dim ldtmFechaGracia As Date = adtmFechaVence
        If ObjDiasGraciaShr.BlnEsValido Then
            If ObjGraciaFinMesBln.ObjValorPro Then
                ldtmFechaGracia = ClsPanorama.FdtmFecUltimoDiaMes(adtmFechaVence)
            Else
                ldtmFechaGracia = adtmFechaVence.AddDays(ObjDiasGraciaShr.ObjValorPro)
            End If
        End If
        Return ldtmFechaGracia
    End Function
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SVacie()
        MyBase.SVacie()
        McolHistServicio = Nothing
        MdtbHistServicio = Nothing
        McolModulosServicio = Nothing
        MdtmFechaGracia = GCDTMFECHANULA
        MdtmFechaVence = GCDTMFECHANULA
        BlnCreandoAno = False
    End Sub
    Protected Overrides Sub SCreeObj(aobjValorLlave() As Object)
        If CType(EnuPermisosObj And EnuPermisosDef.enuCrear, Boolean) Then
            Try
                If FblnEsCreable(aobjValorLlave) Then
                    EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando
                    SVacie()
                    SInicialiceObj()
                    If HenuTipoObjeto = EnuModoInstanciaObjDef.enuNavegable OrElse
                            HenuTipoObjeto = EnuModoInstanciaObjDef.enuUnico Then
                        If Not IsNothing(DrwRegistroActual) Then
                            DrwRegistroActual = DrwRegistroActual.Table.NewRow
                        Else
                            If Not Me.HenuIdClase = EnuIdClasesPanDef.EnuImportar Then
                                Dim ldtbObj As DataTable =
                                        ClsPanorama.FdtbDataTable(HstrNombreTabla,
                                        {"*"}, {{"", ""}}, "")
                                DrwRegistroActual = ldtbObj.NewRow
                            End If
                        End If
                    End If
                    ObjValorLlave = aobjValorLlave
                End If
            Catch ex As PanDatException
                Throw
            Catch ex As PanLException
                Throw
            Catch ex As Exception
                Throw
            End Try
        End If
    End Sub
    Protected Overrides Sub SInicialiceObj()
        ObjIdCarpeta_ServicioShr.ObjValorPro = GshrIdCarpeta
        ObjIdCentroUtil_ServicioShr.ObjValorPro = GshrIdCentroUtil
        ObjConceptoServicioStr.ObjValorPro = String.Empty
        ObjEsServicioIdBln.ObjValorPro = False
        ObjTarifaIvaDbl.ObjValorPro = 0
        ObjTarifaRetFteDbl.ObjValorPro = 0
        ObjTarifaRetIcaDbl.ObjValorPro = 0
        ObjBaseMinimaReteFuenteDec.ObjValorPro = 0
        ObjBaseMinimaReteIcaDec.ObjValorPro = 0
        ObjCodigoCuentaIvaStr.ObjValorPro = String.Empty
        ObjCodigoCuentaCrStr.ObjValorPro = String.Empty
        ObjCodigoCuentaDbStr.ObjValorPro = String.Empty
        ObjCodigoCuentaDevStr.ObjValorPro = String.Empty
        ObjCodigoCuentaMoraStr.ObjValorPro = String.Empty
        ObjEsAjusteBln.ObjValorPro = False
        ObjIdServicioAjustadoShr.ObjValorPro = 0
        ObjIdTipoTerCtaCrSerByt.ObjValorPro = EnuTipoTerCtaCrServicio.EnuCliente
        ObjDiaFacturaShr.ObjValorPro = 1
        ObjDiasVencimientoShr.ObjValorPro = 9
        ObjDiasGraciaShr.ObjValorPro = 0
        ObjFactAPropYPreAgrBln.ObjValorPro = False
        ObjVenceFinMesBln.ObjValorPro = False
        ObjGraciaFinMesBln.ObjValorPro = False
        ObjEstaActivoServicioBln.ObjValorPro = True
        If IsNothing(ObjMiAno) Then
            SInicialiseServicioPer()
        Else
            SInicialiceServicioAno()
        End If
    End Sub
    Friend Sub SInicialiceServicioAno()
        Dim lshrIdAnoPerAct As Short = GobjParametros.ObjAnoActual.ObjIdAnoShr.ObjValorPro
        ObjIdAno_ServicioShr.ObjValorPro = ObjMiAno.ObjIdAnoShr.ObjValorPro
        ObjNombreServicioStr.ObjValorPro = "Cuota Administración"
        ObjTipoBaseCalculoByt.ObjValorPro = ObjMiAno.ObjTipoCalculoCuotaByt.ObjValorPro
        ObjGraciaFinMesBln.ObjValorPro = False
        ObjPeriodoInicioStr.ObjValorPro = ObjMiAno.ObjIdAnoShr.ToString & "01"
        ObjCantPeriodos_ServicioShr.ObjValorPro = 12
        ObjEsFactProgramableBln.ObjValorPro = True
        ObjGeneraProgramBln.ObjValorPro = True
        ObjIdTipoServicioByt.ObjValorPro = EnuTipoServicio.EnuAnual
        ObjEstaAjustadoBln.ObjValorPro = True
        If GobjParametros.FblnPerActEsDicPrimerAno AndAlso lshrIdAnoPerAct =
                        ObjMiAno.ObjIdAnoShr.ObjValorPro Then
            ObjEstaGenaradaProgramBln.ObjValorPro = True
        Else
            ObjEstaGenaradaProgramBln.ObjValorPro = False
        End If
    End Sub
    Private Sub SInicialiseServicioPer()
        If IsNothing(ObjMiAno) Then
            ObjIdTipoServicioByt.ObjValorPro = EnuTipoServicio.EnuPermanente
            ObjIdAno_ServicioShr.ObjValorPro = 0
            ObjNombreServicioStr.ObjValorPro = String.Empty
            ObjTipoBaseCalculoByt.ObjValorPro = EnuTipoBaseCalculo.None
            ObjEstaGenaradaProgramBln.ObjValorPro = False
            ObjEsFactProgramableBln.ObjValorPro = False
            ObjGeneraProgramBln.ObjValorPro = False
            ObjVenceFinMesBln.ObjValorPro = False
            ObjGraciaFinMesBln.ObjValorPro = False
            ObjPeriodoInicioStr.ObjValorPro = String.Empty
            ObjCantPeriodos_ServicioShr.ObjValorPro = 0
            ObjEstaAjustadoBln.ObjValorPro = False
        Else
            Throw New ErrorInesperadoPanLException("Servicio Pertmanente con Año no Null")
        End If
    End Sub
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        Dim lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            GobjPanDat.SInicialiceTransaccion()
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                SNumereObj()
            Else
                SVerifiqueProgramacion()
            End If
            ClsPanorama.SActualiceCol(ColModulosServicio)
            BlnSincronizando = False
            MyBase.SActualice(ablnExigeRequeridos)
            ObjMiAno?.SAjustePresupuesto()
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
    End Sub
    Friend Overrides Function FblnEsCreable(aobjValorLlave As Object()) As Boolean
        Dim lblnEsCreable = MyBase.FblnEsCreable(aobjValorLlave)
        If ObjMiAno IsNot Nothing Then
            If lblnEsCreable AndAlso EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                If GobjParametros.ColAnos.Count = 1 AndAlso ObjMiAno.ColServiciosAno.Count = 0 Then
                    SLevanteEventoNot("Para crear el Servicio, debe haber antes " &
                        "parametrizado el Año!", "", 0, EnuSeveridadNot.EnuAdvertencia)
                End If
            End If
        End If
        Return lblnEsCreable
    End Function
    Friend Overrides Function FblnEsModificable() As Boolean
        Dim lblnEsModi = MyBase.FblnEsModificable()
        If lblnEsModi AndAlso ObjMiAno IsNot Nothing Then
            lblnEsModi = Not ObjMiAno.ObjEstaCerradoAnoBln.ObjValorPro
        End If
        Return lblnEsModi
    End Function
    Protected Overrides Function FblnSuprimio() As Boolean
        Dim lblnSuprimio = False, lblnNoHayError As Boolean
        Try
            GobjPanDat.SControleProcesoObj(True)
            If FblnEsSuprimible() Then
                GobjPanDat.SInicialiceTransaccion()
                SElimineItemsProgFact()
                lblnSuprimio = ClsPanorama.FblnSuprimioCol(ColModulosServicio)
                If lblnSuprimio Then
                    HblnEsSuprimible = True
                    lblnSuprimio = MyBase.FblnSuprimio()
                End If
                If lblnSuprimio Then
                    GobjPanorama.SRegistreAccionLogApp(HstrNombreClase, "Suprimir Servicio No. " &
                                    ObjIdServicioShr.ToString)
                    GobjPanDat.SConfirmeTransaccion()
                Else
                    GobjPanDat.SAborteTransaccion()
                End If
            End If
            lblnNoHayError = True
        Catch ex As ProveedorBdPanException
            Throw
        Catch ex As ArgumentOutOfRangeException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                GobjPanDat.SControleProcesoObj(False)
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
        Return lblnSuprimio
    End Function
    Friend Overrides Function FblnEsSuprimible() As Boolean
        Dim lblnEsSuprimible = EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando
        If lblnEsSuprimible Then
            lblnEsSuprimible = EnuPermisosObj And EnuPermisosDef.enuSuprimir
        End If
        If lblnEsSuprimible Then
            If ObjIdTipoServicioByt.ObjValorPro = EnuTipoServicio.EnuAnual AndAlso
                    Not ObjEsAjusteBln.ObjValorPro Then
                lblnEsSuprimible = ObjMiAno.FblnEsSuprimible()
            ElseIf ObjEsAjusteBln.ObjValorPro Then
                lblnEsSuprimible = True
            Else
                lblnEsSuprimible = False
            End If
        End If
        Return lblnEsSuprimible
    End Function
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjIdServicioShr.ToString
        End Get
    End Property
#End Region
#Region "Procedimientos del objeto"
    Private Sub SNumereObj()
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            Dim lshrIdServicio As Short
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsIdAno_ServicioShr.SstrNombreCampoBd & " = " & ObjIdAno_ServicioShr.ToString
            lshrIdServicio = ClsPanorama.FobjUltimaIdNumericaObjeto(SstrNombreTabla,
                    ClsIdServicioShr.SstrNombreCampoBd, ObjIdServicioShr.EnuTipoValor, lstrFiltro) + 1
            ObjIdServicioShr.ObjValorPro = lshrIdServicio
            For Each lobjModServicio As ClsModuloServicio In ColModulosServicio
                lobjModServicio.ObjIdServicio_ModuloServicioShr.ObjValorPro = lshrIdServicio
            Next
        End If
    End Sub
    Private Sub SVerifiqueProgramacion()
        If ObjGeneraProgramBln.ObjValorPro OrElse (ObjEsFactProgramableBln.ObjValorPro AndAlso
                Not BlnEsImportado) Then
            If ObjPeriodoInicioStr.BlnCambio OrElse ObjCantPeriodos_ServicioShr.BlnCambio OrElse
                    FblnCambioValor() OrElse ObjGeneraProgramBln.BlnCambio Then
                ObjEstaGenaradaProgramBln.ObjValorPro = False
            End If
        End If
    End Sub
    Friend Sub SApliqueCambiosServicio(ablnEstaGenerado As Boolean, ablnEstaAjustada As Boolean)
        If ObjIdAno_ServicioShr.ObjValorPro > 0 Then
            If Not ObjMiAno.ObjModuloPorServicioBln.ObjValorPro Then
                If BlnEsCuotaAdministracion AndAlso Not ObjEsAjusteBln.ObjValorPro Then
                    If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                        EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                    End If
                    ObjEstaGenaradaProgramBln.ObjValorPro = ablnEstaGenerado
                    ObjEstaAjustadoBln.ObjValorPro = ablnEstaAjustada
                End If
            End If
        Else
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
            End If
            ObjEstaGenaradaProgramBln.ObjValorPro = ablnEstaGenerado
        End If
    End Sub
    Friend Sub SActuEstaGeneradoProgramaFact(ablnEstaGenerado As Boolean)
        Dim lblnEstadoConsultando = EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando
        If lblnEstadoConsultando Then
            EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
        End If
        ObjEstaGenaradaProgramBln.ObjValorPro = ablnEstaGenerado
        If ablnEstaGenerado AndAlso lblnEstadoConsultando Then
            SGenereHistorico()
        End If
        If lblnEstadoConsultando Then
            SActualice(True)
        End If
    End Sub
    Friend Sub SActuVlrModulo(adecValorTotal As Decimal)
        If ObjMiAno.ObjModuloPorServicioBln.ObjValorPro Then
            If ColModulosServicio.Count = 1 Then
                EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                Dim lobjModSer As ClsModuloServicio = ColModulosServicio(1)
                Dim lenuTipoBase As EnuTipoBaseCalculo = ObjTipoBaseCalculoByt.ObjValorPro
                With lobjModSer
                    If .EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                        .EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                    End If
                    Dim ldecValorPartModulo = lobjModSer.FdecValorParticipa(adecValorTotal,
                            lenuTipoBase)
                    .ObjValorPres_ModuloServicioDec.ObjValorPro = ldecValorPartModulo
                End With
                SActualice(True)
            Else
                Throw New ErrorInesperadoPanLException("El Servicio no tiene un Módulo por Servicio!")
            End If
        Else
            Throw New ErrorInesperadoPanLException("El Servicio tiene más de un Modulo")
        End If
    End Sub
    Friend Function FdecValorInicialSector(ashrIdSector As Short) As Decimal
        Dim ldtbSecSer = FdtbSectoresServicio()
        Dim lstrFiltro = ClsIdSector_SectorModuloServicioShr.SstrNombreCampoBd & " = " &
                ashrIdSector.ToString
        Dim ldrwSectoresModulo = ldtbSecSer.Select(lstrFiltro)
        Dim ldecValor = 0D
        If ldrwSectoresModulo.Count > 0 Then
            Dim ldrwSectorModulo = ldrwSectoresModulo(0)
            Dim lstrCampo As String
            lstrCampo = ClsValor_SectorModuloServicioDec.SstrNombreCampoBd
            ldecValor = ClsPanorama.FobjValorCampo(ldrwSectorModulo(lstrCampo), EnuTipoValor.enuDecimal)
        End If
        Return ldecValor
    End Function
    Friend Function FdtbSectoresServicio() As DataTable
        GobjPanDat.SControleProcesoObj(True)
        Dim lstrTablaPri = ClsSector.SstrNombreTabla
        Dim lstrTablaSec = ClsSectorModuloServicio.SstrNombreTabla
        Dim lstrCamposTabPri = {ClsNombreSectorStr.SstrNombreCampoBd}
        Dim lstrCamposTabSec = {ClsIdSector_SectorModuloServicioShr.SstrNombreCampoBd,
                    "SUM(" & ClsValor_SectorModuloServicioDec.SstrNombreCampoBd & ")"}
        Dim lstrCamposPrimRel = {StrCampoCarpeta,
                StrCampoCentroUtil, ClsIdSectorShr.SstrNombreCampoBd}
        Dim lstrCamposRelSec = {StrCampoCarpeta,
                StrCampoCentroUtil,
                ClsIdSector_SectorModuloServicioShr.SstrNombreCampoBd}
        Dim lstrIndice = {{ClsIdSector_SectorModuloServicioShr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion_Pri & " AND " &
                ClsIdAno_SectorModuloServicioShr.SstrNombreCampoBd & " = " &
                ObjIdAno_ServicioShr.ToString & " AND " &
                ClsIdServicio_SectorModuloServicioShr.SstrNombreCampoBd & " = " &
                ObjIdServicioShr.ObjValorPro
        Dim lstrGroupBy() = {ClsIdSector_SectorModuloServicioShr.SstrNombreCampoBd}
        Dim ldtbSecSer As DataTable
        Try
            ldtbSecSer = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCamposTabPri,
                    lstrTablaSec, lstrCamposTabSec, lstrCamposPrimRel, lstrCamposRelSec,
                    lstrIndice, lstrFiltro, lstrGroupBy, True)
        Catch ex As ProveedorBdPanException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        End Try
        GobjPanDat.SControleProcesoObj(False)
        Return ldtbSecSer
    End Function
    Friend Function FblnSectoresAsignadosaModulos()
        Dim lblnAsignados = True
        Dim lobjModuloContr As ClsModuloContribucion
        If ObjEsFactProgramableBln.ObjValorPro AndAlso ObjGeneraProgramBln.ObjValorPro Then
            For Each lobjModulo As ClsModuloServicio In ColModulosServicio
                lobjModuloContr = lobjModulo.ObjMiModuloContribucion
                If lobjModuloContr.ColSectoresModulo.Count = 0 Then
                    lblnAsignados = False
                    Exit For
                End If
            Next
        End If
        Return lblnAsignados
    End Function
    Friend Function FblnSectContrConTotalArea() As Boolean
        Dim lblnSi As Boolean
        For Each lobjModuloSer As ClsModuloServicio In ColModulosServicio
            lblnSi = lobjModuloSer.FblnSectContriConTotalArea
            If Not lblnSi Then Exit For
        Next
        Return lblnSi
    End Function
    Friend Function FblnModuloMeContribuye(ashrIdModulo As Short) As Boolean
        Dim lblnContr = False
        If ColModulosServicio.Count > 0 Then
            lblnContr = ColModulosServicio.Contains(ashrIdModulo.ToString)
        End If
        Return lblnContr
    End Function
    Private Sub SElimineItemsProgFact()
        Dim lstrTabla = ClsItemProgramaFact.SstrNombreTabla
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " And " & ClsIdAno_ServicioShr.SstrNombreCampoBd +
                " = " & ObjIdAno_ServicioShr.ToString & " And " & ClsIdServicioShr.SstrNombreCampoBd +
                " = " & ObjIdServicioShr.ToString
        Dim lstrSql = ClsPanoramaDat.FstrConstruyaExpSqlEliminar(lstrTabla, lstrFiltro)
        GobjPanDat.SEjecuteSentenciaSql(lstrSql)
    End Sub
    Friend Function FblnValorOriginalCero() As Boolean
        Dim ldecVlroriginal = 0D
        For Each lobjModuSErv As ClsModuloServicio In ColModulosServicio
            ldecVlroriginal += lobjModuSErv.ObjValorPres_ModuloServicioDec.ObjValorOriginal
        Next
        Dim lblnVlrOrigCero = ldecVlroriginal = 0
        Return lblnVlrOrigCero
    End Function
    Friend Sub SMarqueAjustadoServicio()
        If ObjEsAjusteBln.ObjValorPro Then
            Dim lshrIdSerAju As Short = ObjIdServicioAjustadoShr.ObjValorPro
            Dim lobjSerAju As ClsServicio = ObjMiAno.ColServiciosAno(lshrIdSerAju)
            If Not lobjSerAju.EnuPermisosObj And EnuPermisosDef.enuModificar Then
                lobjSerAju.EnuPermisosObj += EnuPermisosDef.enuModificar
            End If
            lobjSerAju.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
            lobjSerAju.ObjEstaAjustadoBln.ObjValorPro = True
            lobjSerAju.SActualice(True)
        End If
    End Sub
    Friend Function FstrMyKey() As String
        Return ObjIdAno_ServicioShr.ToString() & "," & ObjIdServicioShr.ToString()
    End Function
    Friend Function FblnCausaMora() As Boolean
        Dim lblnCausa = ObjModoCausaInteresesByt.ObjValorPro > EnuModoCausaMora.EnuNoCausa
        Return lblnCausa
    End Function
#End Region
#Region "Estado del Proceso de Calculo"
    Friend Function FblnEstaCalculadaCuotaAdmin() As Boolean
        Dim lblnEstaCalculada = False
        If ObjIdTipoServicioByt.ObjValorPro = EnuTipoServicio.EnuAnual Then
            lblnEstaCalculada = (ObjMiAno.ObjIdAnoShr.ObjValorPro =
                GobjParametros.ObjAnoActual.ObjIdAnoShr.ObjValorPro)
            If lblnEstaCalculada AndAlso ObjGeneraProgramBln.ObjValorPro Then
                lblnEstaCalculada = ObjEstaGenaradaProgramBln.ObjValorPro
            End If
        End If
        Return lblnEstaCalculada
    End Function
    ''' <summary>
    ''' Indica si las Cuotas de Administración correspondientes a este Servicio deben ser ajustadas
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>Las Cuotas de Administración deben ser ajustadas cuando:
    ''' a) Ya fueron generadas las cuotas definitivas,
    ''' b) El Período actual es igual al Período de Aplicación del Presupuesto Definitivo y
    ''' c) No han sido ajustadas
    ''' </remarks>
    Friend Function FblnDebeAjustarCuotasAdmin() As Boolean
        Dim lblnDebeAjustar = BlnEsCuotaAdministracion AndAlso
                (Not ObjMiAno.ObjModuloPorServicioBln.ObjValorPro) AndAlso
                (Not ObjEsAjusteBln.ObjValorPro)
        If lblnDebeAjustar Then
            lblnDebeAjustar = (ObjMiAno.ObjIdAnoShr.ObjValorPro =
                    GobjParametros.ObjAnoActual.ObjIdAnoShr.ObjValorPro)
        End If
        If lblnDebeAjustar Then
            lblnDebeAjustar = Not ObjEstaAjustadoBln.ObjValorPro
        End If
        Return lblnDebeAjustar
    End Function
#End Region
#Region "Manejo Modulos del Servicio"
    Friend ReadOnly Property ColModulosServicio As Collection
        Get
            If McolModulosServicio Is Nothing Then
                McolModulosServicio = New Collection
                Dim ldtbModSer = FdtbModulosServicio()
                For Each ldrwModuloSer As DataRow In ldtbModSer.Rows
                    Dim lobjModuloServicio As New ClsModuloServicio(Me, ldrwModuloSer)
                    lobjModuloServicio.SLeaValores(True)
                    McolModulosServicio.Add(lobjModuloServicio,
                        lobjModuloServicio.ObjIdModulo_ModuloServicioShr.ToString)
                Next
            End If
            Return McolModulosServicio
        End Get
    End Property
    Friend Function FobjNewModuloSer() As ClsModuloServicio
        Dim ldtbModSer As DataTable = FdtbModulosServicio()
        Dim ldrwNewModSer = ldtbModSer.NewRow
        Dim lobjNewModSer As New ClsModuloServicio(Me, ldrwNewModSer)
        lobjNewModSer.SCreeObj(Nothing)
        lobjNewModSer.ObjIdServicio_ModuloServicioShr.ObjValorPro = ObjIdServicioShr.ObjValorPro
        lobjNewModSer.ObjValorPres_ModuloServicioDec.ObjValorPro = 0
        Return lobjNewModSer
    End Function
    Private Sub SAdicioneModuloSer(aobjModuloServicio As ClsModuloServicio)
        Dim lstrKey = aobjModuloServicio.ObjIdModulo_ModuloServicioShr.ToString
        If McolModulosServicio Is Nothing Then
            McolModulosServicio = New Collection
        End If
        McolModulosServicio.Add(aobjModuloServicio, lstrKey)
    End Sub
    Friend Sub SElimineModuloSer(ashrIdModSer As Short)
        If ColModulosServicio.Contains(ashrIdModSer.ToString) Then
            Dim lblnElimino = False
            Dim lobjModuloSer As ClsModuloServicio = ColModulosServicio(ashrIdModSer.ToString)
            If lobjModuloSer.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                If Not lobjModuloSer.FblnSuprimio() Then
                    SLevanteEventoNot("No fue posible eliminar el Módulo de Contribución del Servicio!",
                        "", 0, EnuSeveridadNot.EnuInformacion)
                Else
                    McolModulosServicio = Nothing
                    lblnElimino = True
                End If
            Else
                If ColModulosServicio.Contains(ashrIdModSer.ToString()) Then
                    ColModulosServicio.Remove(ashrIdModSer.ToString())
                    lblnElimino = True
                End If
            End If
            If lblnElimino Then
                SApliqueCambiosServicio(False, False)
            End If
        End If
    End Sub
    Friend Sub SAdicioneNewModuloSer(aobjNewModuloSer As ClsModuloServicio)
        If EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            aobjNewModuloSer.ObjIdServicio_ModuloServicioShr.ObjValorPro =
                    ObjIdServicioShr.ObjValorPro
            ColModulosServicio.Add(aobjNewModuloSer,
                    aobjNewModuloSer.ObjIdModulo_ModuloServicioShr.ToString())
        End If
    End Sub
    Friend Function FdtbModulosServicio() As DataTable
        Dim lentIdAno = 0, ldtbModulosServicio As DataTable
        If Not IsNothing(ObjIdAno_ServicioShr.ObjValorPro) Then
            lentIdAno = ObjIdAno_ServicioShr.ObjValorPro
        End If
        Dim lstrCamposSelect() As String = {StrCampoCarpeta,
                StrCampoCentroUtil,
                ClsIdAno_ModuloServicioShr.SstrNombreCampoBd,
                ClsIdServicio_ModuloServicioShr.SstrNombreCampoBd,
                ClsIdModulo_ModuloServicioShr.SstrNombreCampoBd,
                "'*' AS NombreModulo", ClsValorPres_ModuloServicioDec.SstrNombreCampoBd}
        Dim lstrIndice(,) As String = {{StrCampoCarpeta, "ASC"},
                {StrCampoCentroUtil, "ASC"},
                {ClsIdAno_ModuloServicioShr.SstrNombreCampoBd, "ASC"},
                {ClsIdServicio_ModuloServicioShr.SstrNombreCampoBd, "ASC"},
                {ClsIdModulo_ModuloServicioShr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdAno_ModuloServicioShr.SstrNombreCampoBd & " = " & lentIdAno & " AND " &
                ClsIdServicio_ModuloServicioShr.SstrNombreCampoBd & " = "
        If Not IsNothing(ObjIdServicioShr.ObjValorPro) Then
            lstrFiltro &= ObjIdServicioShr.ObjValorPro
        Else
            lstrFiltro &= My.Resources.Cero
        End If
        ldtbModulosServicio = ClsPanorama.FdtbDataTable(ClsModuloServicio.SstrNombreTabla,
                lstrCamposSelect, lstrIndice, lstrFiltro)
        SRepuebleNombresModulos(ldtbModulosServicio)
        Return ldtbModulosServicio
    End Function
    ''' <summary>
    ''' Indica si los valores de los Modulos de Servicio que contribuyen al Servicio
    ''' ya fueron ingresados en su totalidad.
    ''' </summary>
    Friend Function FblnVlrsModsIngresados() As Boolean
        Dim lblnValoresModsIngresados = False
        If ObjGeneraProgramBln.ObjValorPro Then
            If ObjIdTipoServicioByt.ObjValorPro = EnuTipoServicio.EnuAnual AndAlso
                    ObjMiAno.ObjModuloPorServicioBln.ObjValorPro Then
                lblnValoresModsIngresados = ObjMiAno.ObjValorPres_AnoDec.ObjValorPro > 0
            Else
                Dim ldecVlrServicio = 0D
                For Each lobjModSer As ClsModuloServicio In ColModulosServicio
                    ldecVlrServicio += lobjModSer.ObjValorPres_ModuloServicioDec.ObjValorPro
                Next
                lblnValoresModsIngresados = ldecVlrServicio > 0
            End If
        End If
        Return lblnValoresModsIngresados
    End Function
    Private Function FblnCambioValor() As Boolean
        Dim lblnCambioValor = False
        If ObjGeneraProgramBln.ObjValorPro Then
            For Each lobjModSer As ClsModuloServicio In ColModulosServicio
                If lobjModSer.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                    lblnCambioValor = lobjModSer.ObjValorPres_ModuloServicioDec.BlnCambio
                    If lblnCambioValor Then Exit For
                End If
            Next
        End If
        Return lblnCambioValor
    End Function
    Friend Sub SRepuebleNombresModulos(adtbModulosSer As DataTable)
        If adtbModulosSer.Rows.Count > 0 Then
            Dim lcolModulos As Collection = GobjParametros.ColModulos
            Dim lobjModulo As ClsModuloContribucion
            For Each ldrwModSer As DataRow In adtbModulosSer.Rows
                lobjModulo = lcolModulos(ldrwModSer(ClsIdModulo_ModuloServicioShr.SstrNombreCampoBd))
                Dim lstrNomMod As String = lobjModulo.ObjNombreModuloStr.ObjValorPro
                ldrwModSer("NombreModulo") = lstrNomMod
            Next
        End If
    End Sub
    Friend Sub SSuprimaModulos()
        If Not ObjGeneraProgramBln.ObjValorPro Then
            For Each lobjModulo As ClsModuloServicio In ColModulosServicio
                Dim lblnSuprimio = lobjModulo.FblnSuprimio()
                If Not lblnSuprimio Then
                    Throw New ErrorInesperadoPanLException("No se suprimio el módulo!")
                End If
            Next
        End If
    End Sub
    Friend Sub SAdicioneModulosServicio(aobjServicioAnoAnt As ClsServicio, adblIncrementoCA As Double)
        If BlnEsCuotaAdministracion AndAlso aobjServicioAnoAnt.BlnEsCuotaAdministracion Then
            Dim lobjNewModServicio As ClsModuloServicio = Nothing
            Dim ldrwNuevoModSer As DataRow = Nothing
            Dim ldtbModSer = aobjServicioAnoAnt.FdtbModulosServicio()
            For Each lobjModuloServicioAnoAnt As ClsModuloServicio In aobjServicioAnoAnt.ColModulosServicio
                ldrwNuevoModSer = ldtbModSer.NewRow
                lobjNewModServicio = New ClsModuloServicio(Me, ldrwNuevoModSer)
                With lobjNewModServicio
                    If Not CType(.EnuPermisosObj And EnuPermisosDef.enuCrear, Boolean) Then
                        .EnuPermisosObj += EnuPermisosDef.enuConsCrear
                    End If
                    .SCreeObj(Nothing)
                    .ObjIdAno_ModuloServicioShr.ObjValorPro = ObjIdAno_ServicioShr.ObjValorPro
                    .ObjIdModulo_ModuloServicioShr.ObjValorPro =
                            lobjModuloServicioAnoAnt.ObjIdModulo_ModuloServicioShr.ObjValorPro
                    .ObjIdServicio_ModuloServicioShr.ObjValorPro =
                            lobjModuloServicioAnoAnt.ObjIdServicio_ModuloServicioShr.ObjValorPro
                    If GblnOK Then
                        .ObjValorPres_ModuloServicioDec.ObjValorPro = Math.Round(
                            lobjModuloServicioAnoAnt.ObjValorPres_ModuloServicioDec.ObjValorPro *
                            (1 + adblIncrementoCA), 0)
                    Else
                        .ObjValorPres_ModuloServicioDec.ObjValorPro = 0D
                    End If
                End With
                SAdicioneModuloSer(lobjNewModServicio)
                ClsCalculosServicios.SActualiceSectoresModuloServicio(lobjNewModServicio)
            Next
        End If
    End Sub
#End Region
#Region "Manejo Historico"
    Friend Sub SGenereHistorico()
        If ObjIdTipoServicioByt.ObjValorPro = EnuTipoServicio.EnuPermanente AndAlso
                ObjGeneraProgramBln.ObjValorPro Then
            Dim lblnCambioPermiso = False
            GobjPanDat.SControleProcesoObj(True)
            SCargueDtbHistServicio()
            Dim ldrwNuevoHistServicio As DataRow = MdtbHistServicio.NewRow
            Dim lobjNuevoHistServicio As New ClsHistServicio(Me, ldrwNuevoHistServicio)
            With lobjNuevoHistServicio
                If Not CType(.EnuPermisosObj And EnuPermisosDef.enuCrear, Boolean) Then
                    .EnuPermisosObj += EnuPermisosDef.enuCrear
                    lblnCambioPermiso = True
                End If
                .SCreeObj(Nothing)
                .ObjFC_HistServicioDtm.ObjValorPro = Now
                .SActualice(True)
                If lblnCambioPermiso Then
                    .EnuPermisosObj -= EnuPermisosDef.enuCrear
                End If
            End With
            lobjNuevoHistServicio.SGenereHistModServicio(ColModulosServicio)
            GobjPanDat.SControleProcesoObj(False)
        End If
    End Sub
    Friend ReadOnly Property ColHistServicios As Collection
        Get
            If IsNothing(McolHistServicio) Then
                McolHistServicio = New Collection
                SCargueDtbHistServicio()
                If Not IsNothing(MdtbHistServicio) Then
                    If MdtbHistServicio.Rows.Count > 0 Then
                        Dim ldrwHistServicios As DataRow() = MdtbHistServicio.Select()
                        For Each ldrwHistSer As DataRow In ldrwHistServicios
                            Dim lobjHistServicio As New ClsHistServicio(Me, ldrwHistSer)
                            lobjHistServicio.SLeaValores(True)
                            McolHistServicio.Add(lobjHistServicio, lobjHistServicio.ObjOrdinal_HistServicioShr.ToString)
                        Next
                    End If
                End If
            End If
            Return McolHistServicio
        End Get
    End Property
    Friend ReadOnly Property DtbHistServicio As DataTable
        Get
            SCargueDtbHistServicio()
            Return MdtbHistServicio
        End Get
    End Property
    Private Sub SCargueDtbHistServicio()
        If ObjIdTipoServicioByt.ObjValorPro = EnuTipoServicio.EnuPermanente AndAlso
                ObjGeneraProgramBln.ObjValorPro Then
            If IsNothing(MdtbHistServicio) Then
                Dim lstrCamposSelect() As String = {"*"}
                Dim lstrIndice(,) As String = {{StrCampoCarpeta, "ASC"},
                                               {StrCampoCentroUtil, "ASC"},
                                               {ClsOrdinal_HistServicioShr.SstrNombreCampoBd, "ASC"}}
                Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion & " AND " &
                        ClsIdServicioShr.SstrNombreCampoBd & " = " & ObjIdServicioShr.ObjValorPro
                MdtbHistServicio = ClsPanorama.FdtbDataTable(ClsHistServicio.SstrNombreTabla,
                        lstrCamposSelect, lstrIndice, lstrFiltro)
            End If
        Else
            MdtbHistServicio = Nothing
        End If
    End Sub
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsBaseMinimaReteFuenteDec
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "BaseMinimaRetefuente"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = "BaseMinRetefuente"
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then
            HobjValorNew = 0
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0,
                Decimal.MaxValue, BlnEsRequerido, HenuTipoValor)
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return Format(HobjValorPro, "c")
        End If
    End Function
End Class
Friend Class ClsBaseMinimaReteIcaDec
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "BaseMinimaReteIca"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = "BaseMinReteIca"
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then
            HobjValorNew = 0
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0,
                Decimal.MaxValue, BlnEsRequerido, HenuTipoValor)
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return Format(HobjValorPro, "c")
        End If
    End Function
End Class
Friend Class ClsCantiPeriodos_ServicioShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "CantidadPeriodos"
    Private ReadOnly MobjPadre As ClsServicio = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "CantidadPeriodos"
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HenuTipoValor = EnuTipoValor.enuShort
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0,
                Short.MaxValue, HblnEsRequerido, EnuTipoValor)
        If Not BlnLeyendoOrigen Then
            Dim ldblValorMin As Double = 1.0
            Dim lshrIdAnoRef = MobjPadre.ObjIdAno_ServicioShr.ObjValorPro
            If GobjParametros.FblnPerActEsDicPrimerAno AndAlso
                    GobjParametros.ObjAnoActual.ObjIdAnoShr.ObjValorPro = lshrIdAnoRef Then
                ldblValorMin = 0
            End If
            HblnEsRequerido = MobjPadre.ObjGeneraProgramBln.ObjValorPro
            HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, ldblValorMin,
                    Short.MaxValue, HblnEsRequerido, EnuTipoValor)
            If MobjPadre.ObjIdTipoServicioByt.ObjValorPro = EnuTipoServicio.EnuAnual Then
                Dim lstrPerioIni As String = MobjPadre.ObjPeriodoInicioStr.ObjValorPro
                Dim lshrMes = 0S
                If MobjPadre.ObjPeriodoInicioStr.BlnEsValido Then
                    lshrMes = CType(Right(lstrPerioIni, 2), Short)
                End If
                If HblnEsValido Then
                    If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                        If GobjParametros.FblnPerActEsDicPrimerAno AndAlso
                                GobjParametros.ObjAnoActual.ObjIdAnoShr.ObjValorPro = lshrIdAnoRef Then
                            HblnEsValido = ldblValorMin = 0
                        Else
                            HblnEsValido = lshrMes + HobjValorNew - 1 <= 12
                        End If
                    Else
                        HblnEsValido = HobjValorNew = HobjValorOriginal
                    End If
                Else
                    If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                        If HobjValorNew = 0 Then
                            HblnEsValido = (lshrMes = 12)
                        End If
                    Else
                        HblnEsValido = (HobjValorNew = HobjValorOriginal)
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
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsCodigoCuentaCrStr
    Inherits ClsCBPropiedad
    Private MstrNombreCuenta As String = String.Empty
    Private Const MCSTRNOMBRECAMPOBD As String = "IdCuentaCr"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "CodigoCuentaCr"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MstrNombreCuenta = String.Empty
        MyBase.SVaciePropiedad()
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 4, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            HblnEsValido = ClsOrionCop.FblnEsValidaCtaContabilidad(HobjValorNew.ToString)
            If HblnEsValido Then
                MstrNombreCuenta = ClsOrionCop.FstrNombreCuentaCon(HobjValorNew.ToString)
            Else
                MstrNombreCuenta = String.Empty
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
Friend Class ClsCodigoCuentaDbStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdCuentaDb"
    Private MstrNombreCuenta As String = String.Empty
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "CodigoCuentaDb"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MstrNombreCuenta = String.Empty
        MyBase.SVaciePropiedad()
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 4, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            HblnEsValido = ClsOrionCop.FblnEsValidaCtaContabilidad(HobjValorNew.ToString)
            If HblnEsValido Then
                MstrNombreCuenta = ClsOrionCop.FstrNombreCuentaCon(HobjValorNew.ToString)
            Else
                MstrNombreCuenta = String.Empty
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
Friend Class ClsCodigoCuentaDevStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdCtaDctoCapital"
    Private MstrNombreCuenta As String = String.Empty
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "CuentaDevolCapital"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MstrNombreCuenta = String.Empty
        MyBase.SVaciePropiedad()
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 4, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            HblnEsValido = ClsOrionCop.FblnEsValidaCtaContabilidad(HobjValorNew.ToString)
            If HblnEsValido Then
                MstrNombreCuenta = ClsOrionCop.FstrNombreCuentaCon(HobjValorNew.ToString)
            Else
                MstrNombreCuenta = String.Empty
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
Friend Class ClsCodigoCuentaIvaStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdCuentaIva"
    Private MstrNombreCuenta As String = String.Empty
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "CodigoCuentaIva"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MstrNombreCuenta = String.Empty
        MyBase.SVaciePropiedad()
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsServicio = ObjPadre
        HblnEsRequerido = (Not lobjPadre.ObjEsExcluidoIvaBln.ObjValorPro) AndAlso
                lobjPadre.ObjTarifaIvaDbl.ObjValorPro > 0
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 4, ShrLongitud, BlnEsRequerido)
        If HblnEsValido AndAlso Not IsNothing(HobjValorNew) AndAlso Not String.IsNullOrEmpty(HobjValorNew) Then
            HblnEsValido = ClsOrionCop.FblnEsValidaCtaContabilidad(HobjValorNew.ToString)
            If HblnEsValido Then
                MstrNombreCuenta = ClsOrionCop.FstrNombreCuentaCon(HobjValorNew.ToString)
            Else
                MstrNombreCuenta = String.Empty
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
Friend Class ClsCodigoCuentaMoraStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdCtaMoraCr"
    Private MstrNombreCuenta As String = String.Empty
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "CuentaIntMora"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MstrNombreCuenta = String.Empty
        MyBase.SVaciePropiedad()
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 4, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            HblnEsValido = ClsOrionCop.FblnEsValidaCtaContabilidad(HobjValorNew.ToString)
            If HblnEsValido Then
                MstrNombreCuenta = ClsOrionCop.FstrNombreCuentaCon(HobjValorNew.ToString)
            Else
                MstrNombreCuenta = String.Empty
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
Friend Class ClsConceptoServicioStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "ConceptoServicio"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Concepto del Servicio"
        HshrLongitud = 50
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 3, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            If HobjValorNew.GetType.Name = "String" Then
                If HobjValorNew.ToString.Length > 50 Then
                    HobjValorNew = HobjValorNew.ToString.Substring(0, 50)
                End If
            End If
            Dim lobjPadre As ClsServicio = ObjPadre
            Dim lstrKey = String.Empty
            HobjValorNew = HobjValorNew.ToString.Trim
            If lobjPadre.ObjIdTipoServicioByt.ObjValorPro = EnuTipoServicio.EnuPermanente Then
                If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
                    HblnEsValido = Not GobjParametros.FblnExisteConcepto(lstrKey, HobjValorNew, True)
                Else
                    lstrKey = "0," & lobjPadre.ObjIdServicioShr.ToString
                    HblnEsValido = Not GobjParametros.FblnExisteConcepto(lstrKey, HobjValorNew, False)
                End If
            End If
            If Not HblnEsValido Then
                HstrMens = "El concepto ya existe y debe ser único para los servicios permanentes!"
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
Friend Class ClsEsAjusteBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "EsAjuste"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "EsAjuste"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsServicio = ObjPadre
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
        If HblnEsValido Then
            If HobjValorNew Then
                HblnEsValido = (lobjPadre.ObjIdTipoServicioByt.ObjValorPro = EnuTipoServicio.EnuAnual)
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If HblnEsValido Then
            If Not BlnLeyendoOrigen Then
                Dim lobjPadre As ClsServicio = ObjPadre
                lobjPadre.ObjIdServicioAjustadoShr.SValide()
                lobjPadre.ObjGeneraProgramBln.SValide()
            End If
        End If
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsEsExcluidoIvaBln
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "EsExcluidoIva"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = "EsExcluidoIva"
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If HblnEsValido Then
            If Not BlnLeyendoOrigen Then
                Dim lobjPadre As ClsServicio = ObjPadre
                lobjPadre.ObjTarifaIvaDbl.SValide()
                lobjPadre.ObjCodigoCuentaIvaStr.SValide()
            End If
        End If
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsEsFactProgramableBln
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "EsFacturacionProgramable"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = "EsFacturacionProgramable"
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsServicio = ObjPadre
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
        If HblnEsValido Then
            If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                If lobjPadre.ObjIdTipoServicioByt.ObjValorPro = EnuTipoServicio.EnuAnual Then
                    HblnEsValido = HobjValorNew
                    If Not HblnEsValido Then
                        HstrMens = "Un Servicio Anual tiene que ser Programable!"
                        SNotifiqueDatInv()
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If HblnEsValido Then
            If Not BlnLeyendoOrigen Then
                Dim lobjPadre As ClsServicio = ObjPadre
                lobjPadre.ObjPeriodoInicioStr.SValide()
                lobjPadre.ObjCantPeriodos_ServicioShr.SValide()
            End If
        End If
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsEsServicioIdBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "EsServicioId"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Es Servicio de Identificacion"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsServicio = ObjPadre
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
        If Not BlnLeyendoOrigen Then
            If HblnEsValido Then
                If HobjValorNew Then
                    HstrMens = String.Empty
                    If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                        HblnEsValido = (lobjPadre.ObjIdTipoServicioByt.ObjValorPro =
                                EnuTipoServicio.EnuPermanente)
                        If Not HblnEsValido Then
                            HstrMens = "El Servicio de Identificación tiene que ser Permanente!"
                        Else
                            If IsNothing(lobjPadre.ObjIdServicioShr.ObjValorPro) Then
                                HblnEsValido = Not GobjParametros.FblnHayServicioId
                            Else
                                HblnEsValido = Not GobjParametros.FblnHayOtroServicioId(
                                        lobjPadre.ObjIdServicioShr.ObjValorPro)
                            End If
                            If Not HblnEsValido Then
                                HstrMens = "Solo se puede definir un solo Servicio de Identificación!"
                            End If
                        End If
                    Else
                        If Not GobjParametros.ObjServicioIdActivoBln.ObjValorPro Then
                            HstrMens = "Para que el Servicio de Identificación funcione " &
                                    "adecuadamente, es necesario activarlo en los " &
                                    "Parámetros de la Copropiedad!"
                        End If
                    End If
                    If Not String.IsNullOrEmpty(HstrMens) Then
                        SNotifiqueDatInv()
                    End If
                End If
            End If
        End If
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return "No"
        Else
            Return ClsPanorama.FstrBuleanoToString(HobjValorPro)
        End If
    End Function
End Class
Friend Class ClsEstaActivoServicioBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Activo"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "EstaActivo"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        HobjValorPro = False
        HobjValorNew = HobjValorPro
        HblnEsValido = False
    End Sub

    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
        HblnEsValido = lblnEsValido
    End Sub

    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsEstaAjustadoBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "EstaAjustado"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Esta Ajustado"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsServicio = ObjPadre
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
        If HblnEsValido Then
            If Not lobjPadre.BlnEsCuotaAdministracion Then
                HblnEsValido = HobjValorNew = False
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsEstaGenaradaProgramBln
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "EstaGeneradaProgramacion"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = "EstaGeneradaProgram"
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsServicio = ObjPadre
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
        If HblnEsValido Then
            If Not lobjPadre.ObjGeneraProgramBln.ObjValorPro AndAlso
                    Not lobjPadre.ObjEsAjusteBln.ObjValorPro Then
                If lobjPadre.ObjTipoBaseCalculoByt.ObjValorPro IsNot Nothing AndAlso
                        lobjPadre.ObjTipoBaseCalculoByt.ObjValorPro <>
                        EnuTipoBaseCalculo.EnuImportadas Then
                    HblnEsValido = HobjValorNew = False
                End If
            End If
        End If
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsGeneraProgramBln
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsServicio = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "EsCalculado"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = "GeneraProgramacion"
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.EnuConsultando Then
                If MobjPadre.ObjIdTipoServicioByt.ObjValorPro = EnuTipoServicio.EnuAnual Then
                    If MobjPadre.ObjTipoBaseCalculoByt.ObjValorPro <>
                            EnuTipoBaseCalculo.EnuImportadas Then
                        HblnEsValido = HobjValorNew = True
                    Else
                        HblnEsValido = If(MobjPadre.ObjEsAjusteBln.ObjValorPro OrElse
                            MobjPadre.ObjTipoBaseCalculoByt.ObjValorPro =
                            EnuTipoBaseCalculo.EnuImportadas, HobjValorNew = False, HobjValorNew = True)
                    End If
                    If Not HblnEsValido AndAlso Not GblnImportando Then
                        HstrMens = "Un Servicio Anual tiene que ser Programado!"
                        SNotifiqueDatInv()
                    End If
                Else
                    If MobjPadre.ObjTipoBaseCalculoByt.ObjValorPro =
                            EnuTipoBaseCalculo.EnuImportadas Then
                        HblnEsValido = HobjValorNew = False
                        If Not HblnEsValido Then
                            HstrMens = "Un Servicio importado no puede generar programación"
                            SNotifiqueDatInv()
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                If Not HobjValorPro Then
                    MobjPadre.ObjEstaGenaradaProgramBln.ObjValorPro = False
                End If
            End If
            MobjPadre.ObjEstaGenaradaProgramBln.SValide()
            MobjPadre.ObjPeriodoInicioStr.SValide()
            MobjPadre.ObjCantPeriodos_ServicioShr.SValide()
            MobjPadre.ObjTipoBaseCalculoByt.SValide()
        End If
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class
Friend Class ClsGraciaFinMes_SerBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "GraciaFinMes"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "GraciaFinMes"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        Dim lobjPadre As ClsServicio = ObjPadre
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
Friend Class ClsIdAno_ServicioShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdAno"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdAñoServicio"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsLlave = True
        HbytPosicionLlave = 2
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, Year(Date.MaxValue),
                BlnEsRequerido, EnuTipoValor)
        If HblnEsValido Then
            Dim lobjPadre As ClsServicio = ObjPadre
            If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                If lobjPadre.ObjIdTipoServicioByt.ObjValorPro = EnuTipoServicio.EnuAnual Then
                    Dim lshrAnoActual As Short =
                            GobjParametros.ObjAnoActual.ObjIdAnoShr.ObjValorPro
                    If ClsOrionCop.BlnProcesoEspecial Then
                        HblnEsValido = HobjValorNew <= lshrAnoActual
                    Else
                        HblnEsValido = HobjValorNew >= lshrAnoActual
                    End If
                Else
                    HblnEsValido = (HobjValorNew = 0)
                End If
            ElseIf lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                HblnEsValido = (HobjValorOriginal = HobjValorNew)
                If Not HblnEsValido Then
                    HstrMens = "No es permitido cambiarle el año a un servicio!"
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
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If HblnEsValido Then
            If Not BlnLeyendoOrigen Then
                Dim lobjPadre As ClsServicio = ObjPadre
                With lobjPadre
                    .ObjPeriodoInicioStr.SValide()
                    .ObjCantPeriodos_ServicioShr.SValide()
                    .ObjIdServicioAjustadoShr.SValide()
                End With
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
Friend Class ClsIdTerceroCtaCrDbl
    Inherits ClsCBPropiedad
    Private MstrNombreTercero As String = String.Empty
    Private MblnExisteTercero As Boolean = False
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTerceroCtaCr"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Id Tercero Cta Cr"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MstrNombreTercero = String.Empty
        MyBase.SVaciePropiedad()
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsServicio = ObjPadre
        HblnEsRequerido = (lobjPadre.ObjIdTipoTerCtaCrSerByt.ObjValorPro =
                EnuTipoTerCtaCrServicio.EnuProveedor)
        If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCDBLMINTERC, GCDBLMAXTERC,
                    BlnEsRequerido, EnuTipoValor)
            If (Not HblnEsRequerido) AndAlso HblnEsValido Then
                HblnEsValido = (HobjValorNew = 0)
                If Not HblnEsValido Then
                    HstrMens = "Si el tipo de Tercero es Cliente, la Id. del Tercero debe ser " &
                            "igual a cero!"
                    SNotifiqueDatInv()
                End If
            End If
        Else
            HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCDBLMINTERC, GCDBLMAXTERC,
                    BlnEsRequerido, EnuTipoValor)
        End If
        MstrNombreTercero = String.Empty
        If HblnEsValido Then
            If Not (IsNothing(HobjValorNew) OrElse String.IsNullOrEmpty(HobjValorNew) OrElse
                    HobjValorNew = 0) Then
                Dim lobjLlave() = {HobjValorNew}
                Dim lobjTercero As New ClsTercero(EnuModoInstanciaObjDef.enuUnico)
                lobjTercero.SAbra(lobjLlave)
                MblnExisteTercero = lobjTercero.BlnExiste
                HblnEsValido = MblnExisteTercero
                If MblnExisteTercero Then
                    MstrNombreTercero = lobjTercero.FstrNombreCompleto()
                Else
                    MstrNombreTercero = String.Empty
                End If
            End If
        End If
    End Sub
    Friend ReadOnly Property BlnExisteTercero As Boolean
        Get
            Return MblnExisteTercero
        End Get
    End Property
    Friend ReadOnly Property StrNombreTercero
        Get
            If BlnEsValido Then
                Return MstrNombreTercero
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
        If Not IsNothing(ObjValorPro) Then
            Return HobjValorPro.ToString
        Else
            Return ""
        End If
    End Function
End Class
Friend Class ClsIdServicioShr
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsServicio = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "IdServicio"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdServicio"
        HenuTipoValor = EnuTipoValor.enuUShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 3
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, 98,
                BlnEsRequerido, EnuTipoValor)
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
            HstrMens = String.Empty
            If HblnEsValido Then
                If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                    If Not BlnLeyendoOrigen Then
                        If Not IsNothing(MobjPadre.ObjIdAno_ServicioShr.ObjValorPro) Then
                            Dim lshrIdAno As Short = MobjPadre.ObjIdAno_ServicioShr.ObjValorPro
                            Dim lobjLlavePrincipal() = {GshrIdCarpeta, GshrIdCentroUtil, lshrIdAno, HobjValorNew}
                            If Not MobjPadre.FblnExisteLlave(lobjLlavePrincipal) Then
                                HstrMens = "La Id. del Servicio ingresada no existe!"
                                HblnEsValido = False
                            End If
                        Else
                            HblnEsValido = False
                        End If
                    End If
                ElseIf MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                    HblnEsValido = (HobjValorOriginal = HobjValorNew)
                    If Not HblnEsValido Then
                        HstrMens = "No es permitido cambiar la Identidad a Objeto alguno!"
                    End If
                End If
            Else
                HstrMens = "La Id. del Servicio ingresada no es válida!"
            End If
            If Not String.IsNullOrEmpty(HstrMens) Then
                SNotifiqueDatInv()
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
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdServicioAjustadoShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdServicioAjustado"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HobjValorNew = 0
        HstrNombre = "IdServicioAjustado"
        HenuTipoValor = EnuTipoValor.enuUShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsServicio = ObjPadre
        HblnEsRequerido = lobjPadre.ObjEsAjusteBln.ObjValorPro
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1,
                Short.MaxValue, BlnEsRequerido)
        If Not BlnLeyendoOrigen Then
            If HblnEsValido Then
                If lobjPadre.ObjEsAjusteBln.ObjValorPro Then
                    If lobjPadre.ObjIdAno_ServicioShr.BlnEsValido Then
                        Dim lshrIdAno As Short = lobjPadre.ObjIdAno_ServicioShr.ObjValorPro
                        If Not IsNothing(lshrIdAno) AndAlso lshrIdAno > 0 Then
                            Dim lobjServicioAjus As ClsServicio =
                                    GobjParametros.ObjServicio(lshrIdAno, HobjValorNew)
                            HblnEsValido = lobjServicioAjus.BlnExiste
                        Else
                            HblnEsValido = False
                        End If
                    Else
                        HblnEsValido = False
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
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsTipoBaseCalculoByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTipoBaseCalculo"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "TipoBaseCalculo"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsServicio = ObjPadre
        HblnEsRequerido = lobjPadre.ObjGeneraProgramBln.ObjValorPro
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew,
                EnuTipoBaseCalculo.EnuCoeficientePro, EnuTipoBaseCalculo.EnuImportadas,
                BlnEsRequerido)
        If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If HblnEsValido AndAlso Not GblnImportando AndAlso HobjValorNew <> HobjValorOriginal Then
                HblnEsValido = HobjValorNew <> EnuTipoBaseCalculo.EnuImportadas
                If Not HblnEsValido Then
                    HblnEsValido = lobjPadre.ObjEsAjusteBln.ObjValorPro
                    If Not HblnEsValido Then
                        HstrMens = "Este tipo de base solo se asigna al momento de importar el Servicio!"
                        SNotifiqueDatInv()
                    End If
                End If
            End If
        End If
        If HblnEsValido AndAlso HobjValorNew = EnuTipoBaseCalculo.EnuCuotaAnterior Then
            HblnEsValido = lobjPadre.ObjIdTipoServicioByt.ObjValorPro = EnuTipoServicio.EnuAnual
            If Not HblnEsValido Then
                HstrMens = "El cálculo con base en el Año anterior, solo es permitido en " &
                    "Servicios del Año!"
            End If
            If HblnEsValido Then
                If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                    HblnEsValido = GobjParametros.ColAnos.Count > 0
                End If
                If Not HblnEsValido Then
                    HstrMens = "El cálculo con base en el Año anterior, no es posible en el " &
                            "primer año de la aplicación!"
                End If
            End If
        End If
        If Not String.IsNullOrEmpty(HstrMens) Then
            SNotifiqueDatInv()
        End If
    End Sub
    Private Sub EPosCambioVlr(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosCambio
        If ObjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            Dim lobjPadre As ClsServicio = ObjPadre
            lobjPadre.ObjEstaGenaradaProgramBln.SValide()
            lobjPadre.ObjGeneraProgramBln.SValide()
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
Friend Class ClsIdTipoServicioByt
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsServicio = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "TipoDeServicio"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = "IdTipoServicio"
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew,
                EnuTipoServicio.EnuAnual, EnuTipoServicio.EnuPermanente, BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                HblnEsValido = (HobjValorOriginal = HobjValorNew)
                If Not HblnEsValido Then
                    HstrMens = "No está permitido cambiar el Tipo de Servicio!"
                    SNotifiqueDatInv()
                End If
            End If
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If HblnEsValido Then
            MobjPadre.ObjIdAno_ServicioShr.SValide()
            MobjPadre.ObjGeneraProgramBln.SValide()
            MobjPadre.ObjEsFactProgramableBln.SValide()
            MobjPadre.ObjPeriodoInicioStr.SValide()
            MobjPadre.ObjCantPeriodos_ServicioShr.SValide()
            MobjPadre.ObjEsAjusteBln.SValide()
            MobjPadre.ObjTipoBaseCalculoByt.SValide()
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
Friend Class ClsIdTipoTerCtaCrSerByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTipoTerCuentaCr"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Tipo Tercero Cuenta Cr"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew,
                EnuTipoTerCtaCrServicio.EnuProveedor, EnuTipoTerCtaCrServicio.EnuCliente,
                BlnEsRequerido)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        Dim lobjPadre As ClsServicio = ObjPadre
        If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            lobjPadre.ObjIdTerceroCtaCrDbl.SValide()
        End If
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return ClsOrionCop.FstrNombreDatoConstanteOri(EnuGrupoConstantesOriDef.EnuTipoTerCtaCrSer,
                        HobjValorPro)
        End If
    End Function
End Class
Friend Class ClsNombreServicioStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Nombre"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "NombreServicio"
        HshrLongitud = 50
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = String.Empty
        If HobjValorNew.GetType.Name = "String" Then
            If HobjValorNew.ToString.Length > 50 Then
                HobjValorNew = HobjValorNew.ToString.Substring(0, 50)
            End If
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 2, ShrLongitud, BlnEsRequerido)
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
Friend Class ClsPeriodoInicioStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PeriodoInicio"
    Private ReadOnly MobjPadre As ClsServicio = Nothing
    Protected Overrides Sub SVaciePropiedad()
        HobjValorNew = GCSTRPERIODONULO
        HobjValorPro = HobjValorNew
    End Sub
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "PeriodoInicio"
        HshrLongitud = 6
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsRequerido = MobjPadre.ObjGeneraProgramBln.ObjValorPro
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, ShrLongitud, ShrLongitud,
                BlnEsRequerido)
        If HblnEsValido Then
            ' Validación del periodo inicial: Para los servicios a los cuales la aplicación les
            ' genera programación: Si es un servicio anual el año del periodo debe ser igual
            ' al año del servicio y el periodo "01"; de lo contrario el periodo debe ser igual
            ' o mayor al periodo actual (año y mes de Periodo Actual).
            ' Excepción: En el proceso de instalación, si se está llevando a cabo en enero, 
            ' el servicio de cuota de administración (anual) del año anterior se crea con el
            ' periodo "000000", lo mismo que el período de aplicación del presupuesto
            ' definitivo.
            If MobjPadre.ObjGeneraProgramBln.ObjValorPro Then
                Dim lstrPeriodo As String = HobjValorNew
                Dim lstrPeriodoActual = GobjParametros.ObjAnoActual.StrIdPeriodoActual
                Dim lshrIdAnoPeriodo As Short = lstrPeriodo.Substring(0, 4)
                If Not HblnEsValido OrElse MobjPadre.EnuEstadoActualizacion <>
                        EnuEstadoObjetoDef.EnuConsultando Then
                    If HobjValorNew <> HobjValorOriginal Then
                        If MobjPadre.ObjIdTipoServicioByt.ObjValorPro =
                                EnuTipoServicio.EnuAnual Then
                            If MobjPadre.ObjEsAjusteBln.ObjValorPro Then
                                HblnEsValido = HobjValorNew = lstrPeriodoActual
                            Else
                                HblnEsValido = HobjValorNew =
                                    MobjPadre.ObjMiAno.ObjIdAnoShr.ToString() & "01"
                            End If
                        Else
                            HblnEsValido = lstrPeriodo >= lstrPeriodoActual
                        End If
                    End If
                End If
            Else
                If BlnEsRequerido Then
                    HblnEsValido = HobjValorNew = GCSTRPERIODONULO
                End If
            End If
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If Not BlnLeyendoOrigen Then
            Dim lobjPadre As ClsServicio = ObjPadre
            If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando AndAlso HblnEsValido Then
                If lobjPadre.ObjIdTipoServicioByt.ObjValorPro = EnuTipoServicio.EnuAnual Then
                    lobjPadre.ObjCantPeriodos_ServicioShr.SValide()
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
Friend Class ClsTarifaIvaDbl
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "TarifaDelIva"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = "TarifaIva"
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        Dim lobjPadre As ClsServicio = ObjPadre
        Dim ldblValorMaximo As Double = 0.5
        If lobjPadre.ObjEsExcluidoIvaBln.ObjValorPro Then
            ldblValorMaximo = 0
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0,
                ldblValorMaximo, BlnEsRequerido, HenuTipoValor)
        If HblnEsValido AndAlso HobjValorNew > 0 Then
            HblnEsValido = GobjParametros.ObjIdCtaReteIvaStr.ToString.Length > 0
            If Not HblnEsValido Then
                HstrMens = "Es necesario definir la Cuenta de Contabilidad para la Retención " &
                        "del IVA!"
                SNotifiqueDatInv()
            End If
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If HblnEsValido Then
            If Not BlnLeyendoOrigen Then
                Dim lobjPadre As ClsServicio = ObjPadre
                lobjPadre.ObjCodigoCuentaIvaStr.SValide()
            End If
        End If
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return Format(HobjValorPro, "p")
        End If
    End Function
End Class
Friend Class ClsTarifaRetFteDbl
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "TarifaRetefuente"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = "TarifaRetefuente"
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0,
                0.5, BlnEsRequerido, HenuTipoValor)
        If HblnEsValido AndAlso HobjValorNew > 0 Then
            HblnEsValido = GobjParametros.ObjIdCtaReteFuenteStr.ToString.Length > 0
            If Not HblnEsValido Then
                HstrMens = "Es necesario definir la Cuenta de Contabilidad para la Retención en la Fuente!"
                SNotifiqueDatInv()
            End If
        End If
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return Format(HobjValorPro, "p")
        End If
    End Function
End Class
Friend Class ClsTarifaRetIcaDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "TarifaReteIca"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "TarifaRetefuente"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0,
                0.01, BlnEsRequerido, HenuTipoValor)
        If HblnEsValido AndAlso HobjValorNew > 0 Then
            HblnEsValido = GobjParametros.ObjIdCtaReteIcaStr.ToString.Length > 0
            If Not HblnEsValido Then
                HstrMens = "Es necesario definir la Cuenta de Contabilidad para la Retención " &
                        "de Industria y Comercio!"
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
            Return Format(HobjValorPro, "p")
        End If
    End Function
End Class
Friend Class ClsDiaFactura_SerShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "DiaFactura"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "DiaDeCorte"
        HenuTipoValor = EnuTipoValor.enuShort
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
Friend Class ClsDiasGracia_SerShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "DiasGracia"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "DiasDeGracia"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsServicio = ObjPadre
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0,
                Short.MaxValue, BlnEsRequerido, HenuTipoValor)
        If HblnEsValido Then
            If lobjPadre.ObjGraciaFinMesBln.ObjValorPro Then
                HblnEsValido = HobjValorNew = 0
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
Friend Class ClsDiasVencimiento_SerShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "DiasVencimiento"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "DiasVencimiento"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsServicio = ObjPadre
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0,
                Short.MaxValue, BlnEsRequerido, HenuTipoValor)
        If HblnEsValido Then
            If lobjPadre.ObjVenceFinMesBln.ObjValorPro Then
                HblnEsValido = HobjValorNew = 0
            Else
                HblnEsValido = HobjValorNew > 0
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
    Private Const MCSTRNOMBRECAMPOBD As String = "FacturarPropConPreAgr"
    Private ReadOnly MobjPadre As ClsServicio
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "FacturarAPropietario"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
        If HblnEsValido AndAlso HobjValorNew Then
            HblnEsValido = MobjPadre.ObjIdTipoServicioByt.ObjValorPro =
                    EnuTipoServicio.EnuPermanente
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
Friend Class ClsVenceFinMes_SerBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "VenceFinMes"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "VenceFinMes"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        Dim lobjPadre As ClsServicio = ObjPadre
        If CType(HobjValorPro, Boolean) Then
            lobjPadre.ObjDiasVencimientoShr.ObjValorPro = 0
        End If
        lobjPadre.ObjDiasVencimientoShr.SValide()
        lobjPadre.ObjGraciaFinMesBln.SValide()
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
Friend Class ClsModoCausaInteresesByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdModoCausaMora"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Modo causa mora"
        HenuTipoValor = EnuTipoValor.EnuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HstrMens = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew,
                EnuModoCausaMora.EnuNoCausa, EnuModoCausaMora.EnuAlReciboCaja, HblnEsRequerido)
        If Not HblnEsValido Then
            HstrMens = "El Modo de causar Intereses de Mora no es Valido!"
            SNotifiqueDatInv()
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        Dim lobjPadre As ClsServicio = ObjPadre
        lobjPadre.ObjDiasGraciaShr.SValide()
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
#End Region
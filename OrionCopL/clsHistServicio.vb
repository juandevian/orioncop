Friend Class ClsHistServicio
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriHistServicios"
    ' Variables de modulo
    Private ReadOnly MobjPadre As clsServicio = Nothing
    Private McolHisModServicio As Collection = Nothing
    Private MdtbHistModServicio As DataTable = Nothing
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Administrador al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwHistServicio">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Friend Sub New(aobjPadre As ClsServicio, adrwHistServicio As DataRow)
        HobjPadre = aobjPadre
        MobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsAnulable = False
        HblnEsSuprimible = False
        HblnEsModificable = False
        '
        DrwRegistroActual = adrwHistServicio
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
            Return EnuIdClasesPanDef.enuHistServicio
        End Get
    End Property

    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Histórico Servicio"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjCantPeriodos_HistServicioShr As New ClsCantiPeriodos_HistServicioShr(Me)
    Friend ReadOnly Property ObjFC_HistServicioDtm As New ClsFechaCreacionDtm(Me)
    Friend ReadOnly Property ObjIdCarpeta_HistServicioShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_HistServicioShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdServicio_HistServicioShr As New ClsIdServicio_HistServicioShr(Me)
    Friend ReadOnly Property ObjNombre_HistServicioStr As New ClsNombre_HistServicioStr(Me)
    Friend ReadOnly Property ObjOrdinal_HistServicioShr As New ClsOrdinal_HistServicioShr(Me)
    Friend ReadOnly Property ObjPeriodoIni_HistServicioStr As New ClsPeriodoIni_HistServicioStr(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjCantPeriodos_HistServicioShr)
                HcolPropiedades.Add(ObjFC_HistServicioDtm)
                HcolPropiedades.Add(ObjIdCarpeta_HistServicioShr)
                HcolPropiedades.Add(ObjIdCentroUtil_HistServicioShr)
                HcolPropiedades.Add(ObjIdServicio_HistServicioShr)
                HcolPropiedades.Add(ObjNombre_HistServicioStr)
                HcolPropiedades.Add(ObjOrdinal_HistServicioShr)
                HcolPropiedades.Add(ObjPeriodoIni_HistServicioStr)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return objOrdinal_HistServicioShr.ToString
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SVacie()
        MyBase.sVacie()
        mcolHisModServicio = Nothing
        mdtbHistModServicio = Nothing
    End Sub

    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        Dim lblnNoHayError = False
        gobjPanDat.sControleProcesoObj(True)
        Try
            If enuEstadoActualizacion = enuEstadoObjetoDef.enuCreando Then
                ObjCantPeriodos_HistServicioShr.ObjValorPro = MobjPadre.ObjCantPeriodos_ServicioShr.ObjValorPro
                ObjIdCarpeta_HistServicioShr.ObjValorPro = GshrIdCarpeta
                ObjIdCentroUtil_HistServicioShr.ObjValorPro = GshrIdCentroUtil
                ObjIdServicio_HistServicioShr.ObjValorPro = MobjPadre.ObjIdServicioShr.ObjValorPro
                ObjNombre_HistServicioStr.ObjValorPro = MobjPadre.ObjNombreServicioStr.ObjValorPro
                ObjPeriodoIni_HistServicioStr.ObjValorPro = MobjPadre.ObjPeriodoInicioStr.ObjValorPro
                SNumereObj()
                MyBase.sActualice(ablnExigeRequeridos)
                lblnNoHayError = True
            Else
                Throw New ErrorInesperadoPanLException("Estado del Objeto HistServicio no esperado!")
            End If
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
                gobjPanDat.sControleProcesoObj(False)
            Else
                gobjPanDat.sControleProcesoObj(False, True)
            End If
        End Try
    End Sub
#End Region
#Region "Procedimientos del objeto"
    Private Sub SNumereObj()
        If enuEstadoActualizacion = enuEstadoObjetoDef.enuCreando Then
            Dim lshrOrdinal As Short
            lshrOrdinal = clsPanorama.fobjUltimaIdNumericaObjeto(sstrNombreTabla,
                    clsOrdinal_HistServicioShr.sstrNombreCampoBd,
                    objOrdinal_HistServicioShr.enuTipoValor,
                    clsOrionCop.strFiltroUbicacion) + 1
            objOrdinal_HistServicioShr.objValorPro = lshrOrdinal
        End If
    End Sub
#End Region
#Region "Manejo HistModServicio"
    Friend Sub SGenereHistModServicio(acolModulosServicio As Collection)
        If Not IsNothing(acolModulosServicio) AndAlso acolModulosServicio.Count > 0 Then
            sCargueDtbHistModServicio()
            Dim lblnCambioPermiso = False
            For Each lobjModuloServicio As clsModuloServicio In acolModulosServicio
                Dim ldrwNewHistModSer As DataRow = mdtbHistModServicio.NewRow
                Dim lobjHistModSer As New ClsHistModServicio(Me, ldrwNewHistModSer)
                With lobjHistModSer
                    If Not CType(.EnuPermisosObj And EnuPermisosDef.enuCrear, Boolean) Then
                        .EnuPermisosObj += EnuPermisosDef.enuCrear
                        lblnCambioPermiso = True
                    End If
                    .SCreeObj(Nothing)
                    .objIdModulo_HistModuloServicioShr.objValorPro =
                            lobjModuloServicio.objIdModulo_ModuloServicioShr.objValorPro
                    .objValor_HistModServicioDec.objValorPro =
                            lobjModuloServicio.ObjValorPres_ModuloServicioDec.objValorPro
                    .sActualice(True)
                    If lblnCambioPermiso Then
                        .EnuPermisosObj -= EnuPermisosDef.enuCrear
                    End If
                End With
            Next
        Else
            Throw New ValorArgumentoInvalidoException("El Servicio debe tener una colección de ModuloServicio")
        End If
    End Sub

    Friend ReadOnly Property ColHistModServicios As Collection
        Get
            If enuEstadoActualizacion = enuEstadoObjetoDef.enuConsultando Then
                If IsNothing(mcolHisModServicio) Then
                    mcolHisModServicio = New Collection
                    sCargueDtbHistModServicio()
                    If Not IsNothing(mdtbHistModServicio) AndAlso mdtbHistModServicio.Rows.Count > 0 Then
                        Dim ldrwHistModulosServicios() As DataRow = mdtbHistModServicio.Select
                        For Each ldrwHisModSer As DataRow In ldrwHistModulosServicios
                            Dim lobjHisModSer As New clsHistModServicio(Me, ldrwHisModSer)
                            lobjHisModSer.sLeaValores(True)
                            mcolHisModServicio.Add(lobjHisModSer,
                                    lobjHisModSer.objOrdinal_HistModServicioShr.ToString)
                        Next
                    End If
                End If
            End If
            Return mcolHisModServicio
        End Get
    End Property

    Friend ReadOnly Property DtbHistModServicio As DataTable
        Get
            sCargueDtbHistModServicio()
            Return mdtbHistModServicio
        End Get
    End Property

    Private Sub SCargueDtbHistModServicio()
        If IsNothing(mdtbHistModServicio) Then
            Dim lstrCamposSelect As String() = {"*", "' ' AS Nombre"}
            Dim lstrIndice = {{StrCampoCarpeta, "ASC"},
                              {StrCampoCentroUtil, "ASC"},
                              {clsOrdinal_HistModServicioShr.sstrNombreCampoBd, "ASC"}}
            Dim lstrFiltro = clsOrionCop.strFiltroUbicacion & " AND " &
                    clsOrdinal_HServicioShr.sstrNombreCampoBd &
                    " = " & objOrdinal_HistServicioShr.ToString
            mdtbHistModServicio = clsPanorama.fdtbDataTable(clsHistModServicio.sstrNombreTabla,
                    lstrCamposSelect, lstrIndice, lstrFiltro)
            sComplementeNombreDtb()
        End If
    End Sub

    Private Sub SComplementeNombreDtb()
        Dim ldrwHistModulosServicios() As DataRow = mdtbHistModServicio.Select
        For Each ldrwHisModSer As DataRow In ldrwHistModulosServicios
            Dim lshrIdModulo As Short = ClsPanorama.FobjValorCampo(ldrwHisModSer(
                        ClsIdModulo_HistModuloServicioShr.SstrNombreCampoBd), EnuTipoValor.enuShort)
            Dim lobjModuloContr As ClsModuloContribucion = GobjParametros.ColModulos(lshrIdModulo.ToString)
            Dim lstrNombre = lobjModuloContr.ObjNombreModuloStr.ToString
            ldrwHisModSer("Nombre") = lstrNombre
        Next
    End Sub
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsCantiPeriodos_HistServicioShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "CantidadPeriodos"

    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "CantidadPeriodosHist"
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HenuTipoValor = EnuTipoValor.enuShort
    End Sub

    Public Overrides Sub SValide()
        Dim lobjPadre As ClsHistServicio = ObjPadre
        Dim lobjAbuelo As ClsServicio = lobjPadre.ObjPadre
        Dim lblnEsValido As Boolean = (HobjValorNew = lobjAbuelo.ObjCantPeriodos_ServicioShr.ObjValorPro)
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

Friend Class ClsIdServicio_HistServicioShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdServicio"

    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdServicioHist"
        HenuTipoValor = EnuTipoValor.enuUShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub

    Public Overrides Sub SValide()
        Dim lobjPadre As ClsHistServicio = ObjPadre
        Dim lobjAbuelo As ClsServicio = lobjPadre.ObjPadre
        Dim lblnEsValido As Boolean = (HobjValorNew = lobjAbuelo.ObjIdServicioShr.ObjValorPro)
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

Friend Class ClsNombre_HistServicioStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Nombre"

    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "NombreServicioHist"
        HshrLongitud = 50
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub

    Public Overrides Sub SValide()
        Dim lobjPadre As ClsHistServicio = ObjPadre
        Dim lobjAbuelo As ClsServicio = lobjPadre.ObjPadre
        Dim lblnEsValido As Boolean = (HobjValorNew = lobjAbuelo.ObjNombreServicioStr.ObjValorPro)
        HblnEsValido = lblnEsValido
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

Friend Class ClsOrdinal_HistServicioShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Ordinal"

    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Ordinal"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 2
    End Sub

    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, Short.MinValue,
                Short.MaxValue, BlnEsRequerido, EnuTipoValor)
        If ObjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
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

Friend Class ClsPeriodoIni_HistServicioStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PeriodoInicio"

    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "PeriodoInicio"
        HshrLongitud = 6
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        hblnEsRequerido = True
    End Sub

    Public Overrides Sub SValide()
        Dim lobjPadre As clsHistServicio = objPadre
        Dim lobjAbuelo As clsServicio = lobjPadre.objPadre
        Dim lblnEsValido As Boolean = (hobjValorNew = lobjAbuelo.objPeriodoInicioStr.objValorPro)
        hblnEsValido = lblnEsValido
    End Sub

    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property

    Public Overrides Function ToString() As String
        If IsNothing(hobjValorPro) Then
            Return ""
        Else
            Return hobjValorPro.ToString
        End If
    End Function
End Class
#End Region

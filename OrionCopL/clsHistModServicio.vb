Friend Class ClsHistModServicio
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriHistModulosServicios"
    ' Variables de modulo
    Private ReadOnly MobjPadre As clsHistServicio = Nothing
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Administrador al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwHistModServicio">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Friend Sub New(aobjPadre As ClsHistServicio, adrwHistModServicio As DataRow)
        HobjPadre = aobjPadre
        MobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsAnulable = False
        HblnEsSuprimible = False
        HblnEsModificable = False
        '
        DrwRegistroActual = adrwHistModServicio
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
            Return EnuIdClasesPanDef.enuHistModServicio
        End Get
    End Property

    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Histórico Módulo Servicio"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjIdCarpeta_HistModServicioShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_HistModServicioShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdModulo_HistModuloServicioShr As New ClsIdModulo_HistModuloServicioShr(Me)
    Friend ReadOnly Property ObjIdServicio_HistModuloServicioShr As New ClsIdServicio_HistModuloServicioShr(Me)
    Friend ReadOnly Property ObjOrdinal_HistModServicioShr As New ClsOrdinal_HistModServicioShr(Me)
    Friend ReadOnly Property ObjOrdinal_HServicioShr As New ClsOrdinal_HServicioShr(Me)
    Friend ReadOnly Property ObjValor_HistModServicioDec As New ClsValor_HistModServicioDec(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjIdCarpeta_HistModServicioShr)
                HcolPropiedades.Add(ObjIdCentroUtil_HistModServicioShr)
                HcolPropiedades.Add(ObjIdModulo_HistModuloServicioShr)
                HcolPropiedades.Add(ObjIdServicio_HistModuloServicioShr)
                HcolPropiedades.Add(ObjOrdinal_HistModServicioShr)
                HcolPropiedades.Add(ObjOrdinal_HServicioShr)
                HcolPropiedades.Add(ObjValor_HistModServicioDec)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    '
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        gobjPanDat.sControleProcesoObj(True)
        Try
            If enuEstadoActualizacion = enuEstadoObjetoDef.enuCreando Then
                sNumereObj()
                ObjIdCarpeta_HistModServicioShr.ObjValorPro = GshrIdCarpeta
                ObjIdCentroUtil_HistModServicioShr.ObjValorPro = GshrIdCentroUtil
                ObjIdServicio_HistModuloServicioShr.ObjValorPro =
                        MobjPadre.ObjIdServicio_HistServicioShr.ObjValorPro
                ObjOrdinal_HServicioShr.ObjValorPro = MobjPadre.ObjOrdinal_HistServicioShr.ObjValorPro
                MyBase.sActualice(ablnExigeRequeridos)
            Else
                Throw New ErrorInesperadoPanLException("Estado del Objeto HistServicio no esperado!")
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
            gobjPanDat.sControleProcesoObj(False)
        End Try
    End Sub

    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjOrdinal_HistModServicioShr.ToString
        End Get
    End Property
#End Region
#Region "Procedimientos del objeto"
    Private Sub SNumereObj()
        If enuEstadoActualizacion = enuEstadoObjetoDef.enuCreando Then
            Dim lshrOrdinal As Short
            lshrOrdinal = clsPanorama.fobjUltimaIdNumericaObjeto(sstrNombreTabla,
                    clsOrdinal_HistModServicioShr.sstrNombreCampoBd,
                    objOrdinal_HistModServicioShr.enuTipoValor,
                    clsOrionCop.strFiltroUbicacion) + 1
            ObjOrdinal_HistModServicioShr.ObjValorPro = lshrOrdinal
        End If
    End Sub
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsIdModulo_HistModuloServicioShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdModuloContribucion"
    Public Sub New(aobjPadre As clsCBObjetoPan)
        MyBase.New(aobjPadre)
        hstrNombre = "IdModuloContribucionHist"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue,
                BlnEsRequerido, EnuTipoValor.enuShort)
        If lblnEsValido Then
            Dim lcolModulos As Collection = GobjParametros.ColModulos
            lblnEsValido = lcolModulos.Contains(HobjValorNew.ToString)
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
Friend Class ClsIdServicio_HistModuloServicioShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdServicio"

    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdServicio_ModSerHist"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub

    Public Overrides Sub SValide()
        Dim lobjPadre As ClsHistModServicio = ObjPadre
        Dim lobjAbuelo As ClsHistServicio = lobjPadre.ObjPadre
        Dim lblnEsValido As Boolean = (HobjValorNew = lobjAbuelo.ObjIdServicio_HistServicioShr.ObjValorPro)
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
Friend Class ClsOrdinal_HistModServicioShr
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
Friend Class ClsOrdinal_HServicioShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "OrdinalHistServicio"

    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "OrdinalHistServicio"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub

    Public Overrides Sub SValide()
        Dim lobjPadre As ClsHistModServicio = ObjPadre
        Dim lobjAbuelo As ClsHistServicio = lobjPadre.ObjPadre
        Dim lblnEsValido As Boolean = (HobjValorNew = lobjAbuelo.ObjOrdinal_HistServicioShr.ObjValorPro)
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
Friend Class ClsValor_HistModServicioDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Valor"

    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "ValorInicicialHistModSer"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        hblnEsRequerido = True
    End Sub

    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, Decimal.MaxValue,
                    BlnEsRequerido, EnuTipoValor.enuDecimal)
    End Sub

    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property

    Public Overrides Function ToString() As String
        If IsNothing(objValorPro) Then
            Return ""
        Else
            Return Format(hobjValorPro, "c")
        End If
    End Function
End Class
#End Region

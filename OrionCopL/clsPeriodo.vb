Friend Class ClsPeriodo
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriPeriodos"
    ' Variables de modulo
#End Region

#Region "Constructores"
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Administrador al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwObjeto">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Friend Sub New(aobjPadre As ClsAno, adrwObjeto As DataRow)
        HobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsAnulable = False
        '
        DrwRegistroActual = adrwObjeto
        DtbTablaColeccion = DrwRegistroActual.Table
        HenuTipoPermiso = EnuPermisosDef.enuHeredado
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
            Return EnuIdClasesPanDef.enuPeriodo
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Período"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjEstaCerradoPeriodoBln As New ClsEstaCerradoPeriodoBln(Me)
    Friend ReadOnly Property ObjFechaFacturacionPeriodoDtm As New ClsFechaFacturacionPeriodoDtm(Me)
    Friend ReadOnly Property ObjIdAnoPeriodoShr As New ClsIdAnoPeriodoShr(Me)
    Friend ReadOnly Property ObjIdCarpetaPeriodoShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtilPeriodoShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdPeriodoShr As New ClsIdPeriodoShr(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjEstaCerradoPeriodoBln)
                HcolPropiedades.Add(ObjFechaFacturacionPeriodoDtm)
                HcolPropiedades.Add(ObjIdAnoPeriodoShr)
                HcolPropiedades.Add(ObjIdCarpetaPeriodoShr)
                HcolPropiedades.Add(ObjIdCentroUtilPeriodoShr)
                HcolPropiedades.Add(ObjIdPeriodoShr)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    Friend ReadOnly Property DtmFechaInicioPeriodo As Date
        Get
            Dim lobjAno As clsAno = objPadre
            Return DateSerial(lobjAno.objIdAnoShr.objValorPro, objIdPeriodoShr.objValorPro, 1)
        End Get
    End Property
    Friend ReadOnly Property DtmFechaFinPeriodo As Date
        Get
            Dim lobjAno As ClsAno = ObjPadre
            Dim lintUltimoDiaMes = Date.DaysInMonth(lobjAno.ObjIdAnoShr.ObjValorPro, ObjIdPeriodoShr.ObjValorPro)
            Return DateSerial(lobjAno.ObjIdAnoShr.ObjValorPro, ObjIdPeriodoShr.ObjValorPro, lintUltimoDiaMes)
        End Get
    End Property
    Friend ReadOnly Property BlnPeriodoFacturado As Boolean
        Get
            Return ObjFechaFacturacionPeriodoDtm.ObjValorPro <> GCDTMFECHANULA
        End Get
    End Property
    Friend ReadOnly Property StrPeriodo As String
        Get
            Return ObjIdAnoPeriodoShr.ToString & ObjIdPeriodoShr.ToString
        End Get
    End Property
#End Region
#End Region

#Region "Procedimientos y funciones invalidantes"
    Friend Overrides Function FblnEsSuprimible() As Boolean
        Dim lobjAno As ClsAno = HobjPadre
        Return lobjAno.FblnEsSuprimible()
    End Function
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjIdPeriodoShr.ToString
        End Get
    End Property
#End Region

#Region "Procedimientos del objeto"
    '
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsEstaCerradoPeriodoBln
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "EstaCerradoPeríodo"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = "EstaCerrado"
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

Friend Class ClsFechaFacturacionPeriodoDtm
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsPeriodo = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "FechaFacturacion"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = "FechaFacturacion"
        HblnRegistrarLogCambio = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        HobjValorPro = GCDTMFECHANULA
        HobjValorNew = HobjValorPro
        HblnEsValido = False
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = True
        If Not BlnLeyendoOrigen Then
            Dim lobjPadre As ClsPeriodo = ObjPadre
            Dim ldtmFechaMinima = lobjPadre.DtmFechaInicioPeriodo
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                ldtmFechaMinima = GCDTMFECHANULA
            End If
            Dim ldtmFechaMaxima = Date.Today
            HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMinima, ldtmFechaMaxima,
                        BlnEsRequerido)
        End If
    End Sub
    Public Overrides Function ToString() As String
        Return HobjValorPro.ToString
    End Function
End Class

Friend Class ClsIdAnoPeriodoShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdAno"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdAñoPeríodo"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 2
    End Sub
    Public Overrides Sub SValide()
        Dim lobjValorIng = HobjValorNew
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 2000, Year(Date.MaxValue),
                            BlnEsRequerido, EnuTipoValor)
        If HblnEsValido Then
            If Not BlnLeyendoOrigen Then
                Dim lobjPadre As ClsPeriodo = ObjPadre
                Dim lobjAbuelo As ClsAno = lobjPadre.ObjPadre
                HblnEsValido = (HobjValorNew = lobjAbuelo.ObjIdAnoShr.ObjValorPro)
            End If
        End If
        If Not HblnEsValido Then
            HstrMens = "La Id. del Año ingresada, '" &
                    lobjValorIng.ToString & "',  no es válida!"
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

Friend Class ClsIdPeriodoShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdPeriodo"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdPeriodo"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HstrOrdenIndice = "ASC"
        HbytPosicionLlave = 3
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, 12,
                BlnEsRequerido, EnuTipoValor)
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
            Return Right("00" & HobjValorPro.ToString, 2)
        End If
    End Function
End Class
#End Region
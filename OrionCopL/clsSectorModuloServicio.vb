Friend Class ClsSectorModuloServicio
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriSectoresModuloServicio"
    ' Variables de modulo
    Private ReadOnly MobjPadre As clsModuloServicio = Nothing
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto ModuloServicio al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwObjeto">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Friend Sub New(aobjPadre As ClsModuloServicio, adrwObjeto As DataRow)
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
            Return EnuIdClasesPanDef.enuSectorModuloServicio
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Sector Módulo Servicio"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjIdAno_SectorModuloServicioShr As New ClsIdAno_SectorModuloServicioShr(Me)
    Friend ReadOnly Property ObjIdCarpeta_SectorModuloServicioShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_SectorModuloServicioShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdModulo_SectorModuloServicioShr As New ClsIdModulo_SectorModuloServicioShr(Me)
    Friend ReadOnly Property ObjIdSector_SectorModuloServicioShr As New ClsIdSector_SectorModuloServicioShr(Me)
    Friend ReadOnly Property ObjIdServicio_SectorModuloServicioShr As New ClsIdServicio_SectorModuloServicioShr(Me)
    Friend ReadOnly Property ObjValor_SectorModuloServicioDec As New ClsValor_SectorModuloServicioDec(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjIdAno_SectorModuloServicioShr)
                HcolPropiedades.Add(ObjIdCarpeta_SectorModuloServicioShr)
                HcolPropiedades.Add(ObjIdCentroUtil_SectorModuloServicioShr)
                HcolPropiedades.Add(ObjIdModulo_SectorModuloServicioShr)
                HcolPropiedades.Add(ObjIdSector_SectorModuloServicioShr)
                HcolPropiedades.Add(ObjIdServicio_SectorModuloServicioShr)
                HcolPropiedades.Add(ObjValor_SectorModuloServicioDec)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        If enuEstadoActualizacion = enuEstadoObjetoDef.enuCreando Then
            gobjPanDat.sControleProcesoObj(True)
            Try
                With MobjPadre
                    ObjIdAno_SectorModuloServicioShr.ObjValorPro = .ObjIdAno_ModuloServicioShr.ObjValorPro
                    ObjIdCarpeta_SectorModuloServicioShr.ObjValorPro = GshrIdCarpeta
                    ObjIdCentroUtil_SectorModuloServicioShr.ObjValorPro = GshrIdCentroUtil
                    ObjIdModulo_SectorModuloServicioShr.ObjValorPro = .ObjIdModulo_ModuloServicioShr.ObjValorPro
                    ObjIdServicio_SectorModuloServicioShr.ObjValorPro = .ObjIdServicio_ModuloServicioShr.ObjValorPro
                End With
                MyBase.sActualice(ablnExigeRequeridos)
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
        Else
            MyBase.sActualice(ablnExigeRequeridos)
        End If
    End Sub
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsIdAno_SectorModuloServicioShr
    Inherits clsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdAno"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdAno_SectorModuloServicio"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 2
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0,
                GCSHRANOMAXIMO, BlnEsRequerido)
        If Not BlnLeyendoOrigen Then
            If lblnEsValido Then
                Dim lobjPadre As ClsSectorModuloServicio = ObjPadre
                Dim lobjAbuelo As ClsModuloServicio = lobjPadre.ObjPadre
                lblnEsValido = (HobjValorNew = lobjAbuelo.ObjIdAno_ModuloServicioShr.ObjValorPro)
            End If
        End If
        If Not lblnEsValido Then
            Throw New ErrorInesperadoPanLException
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
Friend Class ClsIdModulo_SectorModuloServicioShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdModuloContribucion"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdModulo_SectorModuloServicio"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 4
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1,
                Short.MaxValue, BlnEsRequerido)
        If Not BlnLeyendoOrigen Then
            If lblnEsValido Then
                Dim lobjPadre As ClsSectorModuloServicio = ObjPadre
                Dim lobjAbuelo As ClsModuloServicio = lobjPadre.ObjPadre
                lblnEsValido = (HobjValorNew = lobjAbuelo.ObjIdModulo_ModuloServicioShr.ObjValorPro)
            End If
        End If
        If Not lblnEsValido Then
            Throw New ErrorInesperadoPanLException
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
Friend Class ClsIdSector_SectorModuloServicioShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdSector"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdSectorContribuyente"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 5
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue, HblnEsRequerido,
                    EnuTipoValor.enuShort)
        If Not BlnLeyendoOrigen Then
            If lblnEsValido Then
                Dim lobjPadre As ClsSectorModuloServicio = ObjPadre
                Dim lobjAbuelo As ClsModuloServicio = lobjPadre.ObjPadre
                Dim lcolSectores As Collection = GobjParametros.ColSectores
                lblnEsValido = lcolSectores.Contains(HobjValorNew.ToString)
                If lblnEsValido Then
                    lblnEsValido = lobjAbuelo.FblnSectorContribuyeModulo(HobjValorNew)
                End If
                If lblnEsValido AndAlso lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                    With lobjPadre
                        Dim lobjValorLlave() = { .ObjIdCarpeta_SectorModuloServicioShr.ObjValorPro,
                                             .ObjIdCentroUtil_SectorModuloServicioShr.ObjValorPro,
                                             .ObjIdAno_SectorModuloServicioShr.ObjValorPro,
                                             .ObjIdServicio_SectorModuloServicioShr.ObjValorPro,
                                             .ObjIdModulo_SectorModuloServicioShr.ObjValorPro, HobjValorNew}
                        lblnEsValido = Not .FblnExisteLlave(lobjValorLlave)
                    End With
                End If
            End If
        End If
        If Not lblnEsValido Then
            Throw New ErrorInesperadoPanLException
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
Friend Class ClsIdServicio_SectorModuloServicioShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdServicio"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdServicio"
        HenuTipoValor = EnuTipoValor.enuUShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 3
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1,
                Short.MaxValue, BlnEsRequerido)
        If Not BlnLeyendoOrigen Then
            If lblnEsValido Then
                Dim lobjPadre As ClsSectorModuloServicio = ObjPadre
                Dim lobjAbuelo As ClsModuloServicio = lobjPadre.ObjPadre
                lblnEsValido = (HobjValorNew = lobjAbuelo.ObjIdServicio_ModuloServicioShr.ObjValorPro)
            End If
        End If
        If Not lblnEsValido Then
            Throw New ErrorInesperadoPanLException
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
Friend Class ClsValor_SectorModuloServicioDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "ValorInicial"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "ValorInicialAno"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0,
                Decimal.MaxValue, BlnEsRequerido, HenuTipoValor)
        If Not lblnEsValido Then
            Throw New ErrorInesperadoPanLException
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
            Return Format(HobjValorPro, "c")
        End If
    End Function
End Class
#End Region

Friend Class ClsSectorModulo
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriSectoresModulos"
    ' Otras variables de modulo
    Private ReadOnly MobjPadre As clsModuloContribucion = Nothing
    Private MobjSector_SectorModulo As ClsSector = Nothing
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia un objeto Panorama
    ''' </summary>
    ''' <param name="aenuModoInstanciaObj">Indica si se instancia como un objeto navegable o como un Objeto único.</param>
    ''' <remarks>Si se instancia como un objeto navegable, se crea un datatable que contiene las columnas de
    ''' la llave con las llaves de todos los objetos y queda a la espera de que se indique que objeto abrir.
    ''' Si se instancia como un objeto único, queda a la espera de recibir el valor de los campos de la llave 
    ''' para abrir dicho objeto. </remarks>
    Public Sub New(aobjPadre As ClsModuloContribucion, aenuModoInstanciaObj As EnuModoInstanciaObjDef)
        If aobjPadre Is Nothing Then
            Throw New ArgumentNullException(NameOf(aobjPadre))
        End If
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuDeColeccion Then
            Throw New ErrorInesperadoPanLException("Con este Constructor no se puede instanciar un Objeto de Colección!")
        End If
        Dim lstrCamposSelect As String()
        HobjPadre = aobjPadre
        MobjPadre = aobjPadre
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuNavegable Then
            HblnEsAnulable = False
            lstrCamposSelect = {StrCampoCarpeta, StrCampoCentroUtil,
                                ClsIdModulo_SectorModuloShr.SstrNombreCampoBd,
                                ClsIdSector_SectorModuloShr.SstrNombreCampoBd}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsIdModulo_SectorModuloShr.SstrNombreCampoBd & " = " & aobjPadre.ObjIdModuloShr.ObjValorPro
            HcolFiltros.Add(lstrFiltro)
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
    End Sub
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Administrador al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwObjeto">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Friend Sub New(aobjPadre As ClsModuloContribucion, adrwObjeto As DataRow)
        HobjPadre = aobjPadre
        MobjPadre = aobjPadre
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
            Return EnuIdClasesPanDef.enuSectorModulo
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Sector Módulo"
        End Get
    End Property
    Friend Overrides ReadOnly Property HstrNombreObj As String
        Get
            Dim lstrNom = Chr(34) & ObjSector_SectorModulo.ObjNombreSectorStr.ObjValorPro & " - " &
                    MobjPadre.ObjNombreModuloStr.ObjValorPro & Chr(34)
            Return lstrNom
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjIdCarpeta_SectorModuloShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_SectorModuloShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdModulo_SectorModuloShr As New ClsIdModulo_SectorModuloShr(Me)
    Friend ReadOnly Property ObjIdSector_SectorModuloShr As New ClsIdSector_SectorModuloShr(Me)
    Friend ReadOnly Property ObjTasaContribucionDbl As New ClsTasaContribucionDbl(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjIdCarpeta_SectorModuloShr)
                HcolPropiedades.Add(ObjIdModulo_SectorModuloShr)
                HcolPropiedades.Add(ObjIdSector_SectorModuloShr)
                HcolPropiedades.Add(ObjTasaContribucionDbl)
                HcolPropiedades.Add(ObjIdCentroUtil_SectorModuloShr)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    Friend ReadOnly Property ObjSector_SectorModulo As ClsSector
        Get
            If IsNothing(MobjSector_SectorModulo) Then
                If ObjIdSector_SectorModuloShr.BlnEsValido Then
                    Dim lcolSectores As Collection = GobjParametros.ColSectores
                    If lcolSectores.Contains(ObjIdSector_SectorModuloShr.ToString) Then
                        MobjSector_SectorModulo = lcolSectores(ObjIdSector_SectorModuloShr.ToString)
                    End If
                End If
            End If
            Return MobjSector_SectorModulo
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SVacie()
        MyBase.SVacie()
        MobjSector_SectorModulo = Nothing
    End Sub
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            GobjPanDat.SControleProcesoObj(True)
            Try
                ObjIdCarpeta_SectorModuloShr.ObjValorPro = GshrIdCarpeta
                ObjIdCentroUtil_SectorModuloShr.ObjValorPro = GshrIdCentroUtil
                ObjIdModulo_SectorModuloShr.ObjValorPro = MobjPadre.ObjIdModuloShr.ObjValorPro
                MyBase.SActualice(ablnExigeRequeridos)
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
        Else
            MyBase.SActualice(ablnExigeRequeridos)
        End If
    End Sub
#End Region
#Region "Procedimientos del objeto"
    Friend Function FdblBasePartiPonderada(aenuBasePartic As EnuTipoBaseCalculo) As Double
        Dim ldblBasePartSector As Double
        ldblBasePartSector = If(aenuBasePartic = EnuTipoBaseCalculo.EnuCoeficientePro,
                ObjSector_SectorModulo.FdblTotalAreaPrediosSector(),
                ObjSector_SectorModulo.FentCantidadPrediosSector)
        Dim ldblBasePartPon As Double
        Dim ldblFactorPondera = ObjTasaContribucionDbl.ObjValorPro
        ldblBasePartPon = ldblBasePartSector * ldblFactorPondera
        Return ldblBasePartPon
    End Function
    Friend Function FblnParticipaAdmon() As Boolean
        Return MobjPadre.ObjContribuyeCuotaAdminBln.ObjValorPro
    End Function
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsIdModulo_SectorModuloShr
    Inherits clsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdModuloContribucion"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdModuloContribucion"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 2
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1,
                Short.MaxValue, BlnEsRequerido)
        If Not BlnLeyendoOrigen Then
            If HblnEsValido Then
                Dim lobjPadre As ClsSectorModulo = ObjPadre
                Dim lobjAbuelo As ClsModuloContribucion = lobjPadre.ObjPadre
                HblnEsValido = (HobjValorNew = lobjAbuelo.ObjIdModuloShr.ObjValorPro)
            End If
            If Not HblnEsValido Then
                HstrMens = "La Id. del Modulo ingresado no es valida!"
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
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdSector_SectorModuloShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdSector"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdSector_SectorModulo"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 3
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPadre As ClsSectorModulo = ObjPadre
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue, HblnEsRequerido,
                    EnuTipoValor.enuShort)
        If HblnEsValido AndAlso Not BlnLeyendoOrigen Then
            If lobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                If lobjPadre.FblnParticipaAdmon Then
                    If GobjParametros.ObjAnoActual IsNot Nothing AndAlso
                        GobjParametros.ObjAnoActual.ObjModuloPorServicioBln.ObjValorPro Then
                        HblnEsValido = Not ClsOrionCop.FblnSectYaContriAdmin(HobjValorNew)
                        If Not HblnEsValido Then
                            HstrMens = "El Sector ya contribuye con un Módulo que contribuye a " &
                                    "Administración!"
                            SNotifiqueDatInv()
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
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsTasaContribucionDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "TasaContribucion"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "TasaContribucion"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = "TasaContribucion"
        HblnRegistrarLogCambio = True
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, 1, HblnEsRequerido,
                HenuTipoValor)
        If HblnEsValido Then
            HobjValorNew = Math.Round(HobjValorNew, 4)
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
#End Region
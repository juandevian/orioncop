Friend Class ClsProveedorEFac
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriProveedoresEFac"
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
    Public Sub New(aObjPadre As ClsCentroUtilOriCop,
                aenuModoInstanciaObj As EnuModoInstanciaObjDef)
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuDeColeccion Then
            Throw New ErrorInesperadoPanLException("Con este Constructor no se puede instanciar un Objeto de Colección!")
        End If
        Dim lstrCamposSelect As String()
        HobjPadre = aObjPadre
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuNavegable Then
            HcolFiltros.Add(ClsOrionCop.StrFiltroUbicacion)
            lstrCamposSelect = {StrCampoCarpeta, StrCampoCentroUtil,
                   ClsIdProveedorEFacEnt.SstrNombreCampoBd}
        Else
            HenuTipoObjeto = EnuModoInstanciaObjDef.enuUnico
            lstrCamposSelect = {"*"}
        End If
        HblnEsAnulable = False
        HblnEsSuprimible = False
        HenuTipoPermiso = EnuPermisosDef.enuHeredado
        HcolTablas.Add(MCSTRNOMBRETABLA)
        HcolCamposSelect.Add(lstrCamposSelect)
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
            Return EnuIdClasesPanDef.enuProveEFac
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Proveedor EFactura"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjContrasenaAPIEFacStr As New ClsContrasenaAPIEFacStr(Me)
    Friend ReadOnly Property ObjIdCarpeta As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdProveedorEFacEnt As New ClsIdProveedorEFacEnt(Me)
    Friend ReadOnly Property ObjIdUsuarioProvEFacStr As New ClsIdUsuarioProvEFacStr(Me)
    Friend ReadOnly Property ObjSubirFacBln As New ClsSubirFacBln(Me)
    Friend ReadOnly Property ObjURLStr As New ClsURLStr(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjContrasenaAPIEFacStr)
                HcolPropiedades.Add(ObjIdCarpeta)
                HcolPropiedades.Add(ObjIdCentroUtil)
                HcolPropiedades.Add(ObjIdProveedorEFacEnt)
                HcolPropiedades.Add(ObjIdUsuarioProvEFacStr)
                HcolPropiedades.Add(ObjURLStr)
                HcolPropiedades.Add(ObjSubirFacBln)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SInicialiceObj()
        ObjIdCarpeta.ObjValorPro = GshrIdCarpeta
        ObjIdCentroUtil.ObjValorPro = GshrIdCentroUtil
        ObjIdProveedorEFacEnt.ObjValorPro = EnuProveedorEFac.None
        ObjIdUsuarioProvEFacStr.ObjValorPro = String.Empty
        ObjContrasenaAPIEFacStr.ObjValorPro = String.Empty
        ObjURLStr.ObjValorPro = String.Empty
    End Sub
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjIdProveedorEFacEnt.ToString
        End Get
    End Property
#End Region
#Region "Procedimientos del objeto"
    Friend Function FstrContrasena()
        Dim lstrContr = String.Empty
        If ObjContrasenaAPIEFacStr.BlnEsValido Then
            lstrContr = ClsPanorama.FstrContrasena(False, ObjContrasenaAPIEFacStr.ObjValorPro)
        End If
        Return lstrContr
    End Function
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsContrasenaAPIEFacStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "ContrasenaEFac"
    Private ReadOnly MobjPadre As ClsProveedorEFac = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Constrasena Proveedor eFac"
        HshrLongitud = 20
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsRequerido = MobjPadre.ObjIdProveedorEFacEnt.ObjValorPro >
                EnuProveedorEFac.None
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 4,
                        ShrLongitud, BlnEsRequerido)
        If Not HblnEsValido Then
            HstrMens = "La contraseña debe contener entre 4 y 20 Caracteres"
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
Friend Class ClsIdProveedorEFacEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdProveedorEFac"
    Private ReadOnly MobjPadre As ClsProveedorEFac = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Id. Proveedor EFactura"
        MobjPadre = aobjPadre
        HenuTipoValor = EnuTipoValor.enuUInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsLlave = True
        HstrOrdenIndice = "ASC"
        HbytPosicionLlave = 2
    End Sub
    Public Overrides Sub SValide()
        HblnEsRequerido = MobjPadre.ObjIdProveedorEFacEnt.ObjValorPro > EnuProveedorEFac.None
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1,
                Integer.MaxValue, BlnEsRequerido)
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
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        MobjPadre.ObjContrasenaAPIEFacStr.SValide()
    End Sub
End Class
Friend Class ClsIdUsuarioProvEFacStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdUsuarioProEFac"
    Private ReadOnly MobjPadre As ClsProveedorEFac = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Id Usuario del Proveedor"
        HenuTipoValor = EnuTipoValor.enuString
        HshrLongitud = 20
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsRequerido = MobjPadre.ObjIdProveedorEFacEnt.ObjValorPro > EnuProveedorEFac.None
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 2, 20, BlnEsRequerido)
        If HblnEsValido Then
            HobjValorNew = HobjValorNew.ToString.Trim
            HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 2, 20, BlnEsRequerido)
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
Friend Class ClsSubirFacBln
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsProveedorEFac = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "SubirFactura"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = ObjPadre
        HstrNombre = "Subir Facura a FTP"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
        If HblnEsValido AndAlso HobjValorNew Then
            HblnEsValido = (MobjPadre.ObjIdProveedorEFacEnt.ObjValorPro > EnuProveedorEFac.None)
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
Friend Class ClsURLStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "URL"
    Private ReadOnly MobjPadre As ClsProveedorEFac = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "URL del Proveedor"
        HshrLongitud = 100
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsRequerido = MobjPadre.ObjIdProveedorEFacEnt.ObjValorPro > EnuProveedorEFac.None
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 14, ShrLongitud, BlnEsRequerido)
        If HblnEsValido AndAlso BlnEsRequerido Then
            HblnEsValido = (HobjValorNew.ToString.Contains("https://")) AndAlso
                    HobjValorNew.ToString.EndsWith("/")
            If HblnEsValido Then
                HobjValorNew = HobjValorNew.ToString.Trim
                HblnEsValido = (HobjValorNew.ToString.Contains("https://")) AndAlso
                    HobjValorNew.ToString.EndsWith("/")
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
#End Region
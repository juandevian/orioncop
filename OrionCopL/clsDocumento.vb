Friend Class ClsDocumento
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriDocumentosContables"
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
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuDeColeccion Then
            Throw New ErrorInesperadoPanLException("Con este Constructor no se puede instanciar un Objeto de Colección!")
        End If
        Dim lstrCamposSelect As String()
        HobjPadre = Nothing
        HblnEsCreable = False
        HblnEsSuprimible = False
        HblnEsAnulable = False
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuNavegable Then
            lstrCamposSelect = {StrCampoCarpeta, StrCampoCentroUtil,
                                ClsIdDocumentoEnt.SstrNombreCampoBd}
            HcolFiltros.Add(ClsOrionCop.StrFiltroUbicacion)
        Else
            HenuTipoObjeto = EnuModoInstanciaObjDef.enuUnico
            lstrCamposSelect = {"*"}
        End If
        HcolTablas.Add(MCSTRNOMBRETABLA)
        HcolCamposSelect.Add(lstrCamposSelect)
        HenuTipoPermiso = EnuPermisosDef.enuTodos
    End Sub

    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Administrador al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwDocumento">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Public Sub New(aobjPadre As ClsCentroUtilOriCop, adrwDocumento As DataRow)
        HobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsAnulable = False
        HblnEsSuprimible = False
        '
        DrwRegistroActual = adrwDocumento
        DtbTablaColeccion = DrwRegistroActual.Table
        HenuTipoPermiso = EnuPermisosDef.enuTodos
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
            Return EnuIdClasesPanDef.enuDocsCont
        End Get
    End Property

    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Documentos Contables"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjIdDocumentoEnt As New ClsIdDocumentoEnt(Me)
    Friend ReadOnly Property ObjIdCarpeta_DocShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_DocShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjNombre_DocStr As New ClsNombre_DocStr(Me)
    Friend ReadOnly Property ObjNumeroInicial_DocEnt As New ClsNumeroInicial_DocEnt(Me)
    Friend ReadOnly Property ObjPrefijo_DocStr As New ClsPrefijo_DocStr(Me)
    Friend ReadOnly Property ObjTipoDocumentoStr As New ClsTipoDocumentoStr(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjIdDocumentoEnt)
                HcolPropiedades.Add(ObjIdCarpeta_DocShr)
                HcolPropiedades.Add(ObjIdCentroUtil_DocShr)
                HcolPropiedades.Add(ObjNombre_DocStr)
                HcolPropiedades.Add(ObjNumeroInicial_DocEnt)
                HcolPropiedades.Add(ObjPrefijo_DocStr)
                HcolPropiedades.Add(ObjTipoDocumentoStr)
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
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjIdDocumentoEnt.ToString
        End Get
    End Property
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        Dim lblnNoHayError = False, lstrMens = String.Empty
        Dim lblnCambioPrefFac = False
        GobjPanDat.SControleProcesoObj(True)
        Try
            If ObjIdDocumentoEnt.ObjValorPro = 1 Then
                lblnCambioPrefFac = ObjPrefijo_DocStr.BlnCambio
            End If
            MyBase.SActualice(ablnExigeRequeridos)
            If lblnCambioPrefFac Then
                GobjParametros.SActualicePieFactura()
                lstrMens = "El Prefijo de Factura en la Resolucion que habilita Numeración, cambió!"
            End If
            lblnNoHayError = True
        Catch ex As ErrorInesperadoPanDatException
            Throw
        Catch es As ErrorInesperadoPanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As PanLException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                If Not String.IsNullOrEmpty(lstrMens) Then
                    SLevanteEventoNot(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                End If
                GobjPanDat.SControleProcesoObj(False)
            Else
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
#End Region
#Region "Procedimientos del objeto"
    '
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsIdDocumentoEnt
    Inherits clsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdDocumento"
    Private ReadOnly MobjPadre As ClsDocumento = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Id Documento"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 2
    End Sub
    Public Overrides Sub SValide()
        Dim lobjValorIng = HobjValorNew
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue,
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
            Return HobjValorPro.ToString
        End If
    End Function
    Private Sub ClsIdDocumentoEnt_evnPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            MobjPadre.ObjTipoDocumentoStr.SValide()
        End If
    End Sub
End Class
Friend Class ClsNombre_DocStr
    Inherits clsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "NombreDocumento"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Nombre Documento"
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HshrLongitud = 50
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 2, Short.MaxValue,
                HblnEsRequerido)
        If HblnEsValido Then
            HobjValorNew = HobjValorNew.ToString.ToUpper
        Else
            If HobjValorNew.ToString.Length < 2 OrElse HobjValorNew.ToString.Length > 50 Then
                HstrMens = "El Nombre debe tener una longitud entre 2 y 50 caracteres!"
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
Friend Class ClsNumeroInicial_DocEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "NumeracionInicialDoc"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Numeracion Inicial Documento"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, Integer.MaxValue,
                BlnEsRequerido)
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
Friend Class ClsPrefijo_DocStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PrefijoDocumento"
    Private ReadOnly MobjPadre As ClsDocumento = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Prefijo Documento"
        HshrLongitud = 10
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud, BlnEsRequerido)
        If Not HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                If HobjValorNew.ToString().Length > 10 Then
                    HstrMens = "El Prefijo del Documento debe tener una longitud entre 0 y 10 caracteres!"
                    SNotifiqueDatInv()
                End If
            End If
        Else
            HblnEsValido = Not HobjValorNew.ToString().Contains("-")
            If Not HblnEsValido Then
                HstrMens = "El prefijo no puede contener guiones!"
                SNotifiqueDatInv()
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
Friend Class ClsTipoDocumentoStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "TipoDocumentoCont"
    Private ReadOnly MobjPadre As ClsDocumento = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Tipo Documento Contable"
        HshrLongitud = 15
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        If GblnActualizandoApp OrElse MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            HblnEsRequerido = False
        Else
            If MobjPadre.ObjIdDocumentoEnt.ObjValorPro < 9 Then
                HblnEsRequerido = (GobjParametros.ObjTipoInterfazByt.ObjValorPro =
                        EnuTipoInterfazDef.EnuPorDocumento)
            Else
                HblnEsRequerido = (GobjParametros.ObjTipoInterfazByt.ObjValorPro =
                        EnuTipoInterfazDef.EnuPorComprobante)
            End If
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
        If IsNothing(HobjValorNew) Then HobjValorNew = String.Empty
        If Not HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                If HobjValorNew.ToString.Length > 15 OrElse HobjValorNew.ToString.Length < 1 Then
                    HstrMens = "El Tipo del Documento debe tener una longitud entre 1 y 15 caracteres!"
                    SNotifiqueDatInv()
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
#End Region
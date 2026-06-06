Imports System.Drawing
Friend Class ClsCuentaBanco
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriCuentasBanco"
    ' Variables de modulo
    Private MobjImagenQR As ClsImagen = Nothing
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
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.enuNavegable Then
            HblnEsAnulable = False
            lstrCamposSelect = {StrCampoCarpeta, StrCampoCentroUtil,
                                ClsIdCuentaBancoShr.SstrNombreCampoBd}
            HcolFiltros.Add(ClsOrionCop.StrFiltroUbicacion)
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
    Friend Sub New(aobjPadre As Object, adrwObjeto As DataRow)
        HobjPadre = aobjPadre
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
            Return EnuIdClasesPanDef.enuCuentaBanco
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Cuenta Bancaria"
        End Get
    End Property
    Friend Overrides ReadOnly Property HstrNombreObj As String
        Get
            Return Chr(34) & "Nro. " & Chr(34) & ObjNumeroCuentaStr.ObjValorPro & Chr(34)
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjEstaActivaBln As New ClsEstaActivaBln(Me)
    Friend ReadOnly Property ObjIdCarpeta_CuentaBancoShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_CuentaBancoShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdCtaContabilidadStr As New ClsIdCtaContabilidadStr(Me)
    Friend ReadOnly Property ObjIdCuentaBancoShr As New ClsIdCuentaBancoShr(Me)
    Friend ReadOnly Property ObjNombreBancoStr As New ClsNombreBancoStr(Me)
    Friend ReadOnly Property ObjNumeroCuentaStr As New ClsNumeroCuentaStr(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjEstaActivaBln)
                HcolPropiedades.Add(ObjIdCarpeta_CuentaBancoShr)
                HcolPropiedades.Add(ObjIdCentroUtil_CuentaBancoShr)
                HcolPropiedades.Add(ObjIdCtaContabilidadStr)
                HcolPropiedades.Add(ObjIdCuentaBancoShr)
                HcolPropiedades.Add(ObjNombreBancoStr)
                HcolPropiedades.Add(ObjNumeroCuentaStr)
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
        Dim lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                SNumereObj()
            End If
            If MobjImagenQR IsNot Nothing Then
                If MobjImagenQR.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                    Dim lentIdImg = CInt(GshrIdCarpeta.ToString() & GshrIdCentroUtil.ToString &
                        ObjIdCuentaBancoShr.ObjValorPro)
                    MobjImagenQR.ObjIdImagenDbl.ObjValorPro = lentIdImg
                    MobjImagenQR.SActualice(True)
                Else
                    If Not ObjEstaActivaBln.ObjValorPro Then
                        FblnSuprimioQR()
                    End If
                End If
            End If
            MyBase.SActualice(ablnExigeRequeridos)
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
                GobjPanDat.SControleProcesoObj(False)
            Else
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    Friend Overrides Function FblnEsSuprimible() As Boolean
        Dim lblnEsSuprimible = FblnPermitidoSuprimir()
        If lblnEsSuprimible Then
            Dim lstrCondicion As String = " = " & ObjIdCuentaBancoShr.ObjValorPro & " AND " &
                    ClsOrionCop.StrFiltroUbicacion
            lblnEsSuprimible = ClsPanorama.FblnEsEliminableReg({SstrNombreTabla},
                    ObjIdCuentaBancoShr.StrNombreCampoBD, lstrCondicion, True, False)
        End If
        Return lblnEsSuprimible
    End Function
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjIdCuentaBancoShr.ToString
        End Get
    End Property
    Protected Overrides Sub SInicialiceObj()
        ObjEstaActivaBln.ObjValorPro = True
        ObjIdCarpeta_CuentaBancoShr.ObjValorPro = GshrIdCarpeta
        ObjIdCentroUtil_CuentaBancoShr.ObjValorPro = GshrIdCentroUtil
        ObjIdCtaContabilidadStr.ObjValorPro = String.Empty
    End Sub
#End Region
#Region "Procedimientos del objeto"
    Private Sub SNumereObj()
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            Dim lshrIdCuentaBanco As Short
            lshrIdCuentaBanco = ClsPanorama.FobjUltimaIdNumericaObjeto(SstrNombreTabla,
                    ClsIdCuentaBancoShr.SstrNombreCampoBd, ObjIdCuentaBancoShr.EnuTipoValor,
                    ClsOrionCop.StrFiltroUbicacion) + 1
            ObjIdCuentaBancoShr.ObjValorPro = lshrIdCuentaBanco
        End If
    End Sub
#End Region
#Region "Manejo imagen QR"
    Friend Function FmstImagenQR() As MemoryStream
        Dim lmstImagen As MemoryStream = Nothing
        Dim ldtbCodigoQR = ClsOrionCop.FdtbImagenBancoQR(ObjIdCuentaBancoShr.ObjValorPro)
        If ldtbCodigoQR.Rows.Count > 0 Then
            MobjImagenQR = New ClsImagen(Me, ldtbCodigoQR.Rows(0)) With {
                        .EnuPermisosObj = EnuPermisosObj
            }
            MobjImagenQR.SLeaValores(True)
            lmstImagen = MobjImagenQR.ObjPropiedadImagenImg.FmstImagenBinaria
        End If
        Return lmstImagen
    End Function
    Friend Sub SAdicioneImagen(aimgFoto As Image)
        Dim ldtbCodigoQRImg = ClsOrionCop.FdtbImagenBancoQR(0)
        If EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando AndAlso
                ldtbCodigoQRImg.Rows.Count = 0 Then
            If Not IsNothing(aimgFoto) Then
                Dim ldrwNuevaImagen As DataRow = ldtbCodigoQRImg.NewRow
                MobjImagenQR = New ClsImagen(Me, ldrwNuevaImagen) With {
                        .EnuPermisosObj = Me.EnuPermisosObj
                    }
                MobjImagenQR.SCreeObj(Nothing)
                Dim lstrNomCoprop =
                        GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual.ObjNombreCentroUtilStr.ObjValorPro
                With MobjImagenQR
                    .ObjIdCategoriaByt.ObjValorPro = EnuCategoriaImagenDef.enuDocumentos
                    .ObjPropiedadImagenImg.BlnLeyendoOrigen = True
                    .ObjPropiedadImagenImg.ObjValorPro = aimgFoto
                    .ObjFechaDtm.ObjValorPro = Date.Today
                    .ObjDescripcionStr.ObjValorPro = "CódigoQR Copropiedad " & lstrNomCoprop
                End With
            End If
        End If
    End Sub
    Friend Function FblnSuprimioQR() As Boolean
        Dim lblnSuprimio = False
        If MobjImagenQR.FblnPermitidoSuprimir() Then
            Dim ldtbCodigoQR = ClsOrionCop.FdtbImagenBancoQR(ObjIdCuentaBancoShr.ObjValorPro)
            If ldtbCodigoQR.Rows.Count > 0 Then
                Dim lobjImagen As New ClsImagen(Me, ldtbCodigoQR.Rows(0)) With {
                        .EnuPermisosObj = EnuPermisosObj
            }
                lobjImagen.SLeaValores(True)
                lblnSuprimio = lobjImagen.FblnSuprimio()
            End If
        End If
        Return lblnSuprimio
    End Function
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsEstaActivaBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "EstaActiva"
    Friend Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Activa"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
        HblnEsValido = lblnEsValido
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
Friend Class ClsIdCtaContabilidadStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdCtaContabilidad"
    Private MstrNombreCuenta As String = String.Empty
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdCtaContabilidad"
        HshrLongitud = 30
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 4, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            HblnEsValido = ClsOrionCop.FblnEsValidaCtaContabilidad(HobjValorNew.ToString)
            If HblnEsValido Then
                MstrNombreCuenta = ClsOrionCop.FstrNombreCuentaCon(HobjValorNew.ToString)
            Else
                HstrMens = "La Cuenta de Contabilidad ingresada no ha sido creada aún!"
                SNotifiqueDatInv()
            End If
        Else
            If ObjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando AndAlso
                    Not String.IsNullOrEmpty(HobjValorNew) Then
                HstrMens = "La Cuenta de Contabilidad ingresada no es válida!"
                SNotifiqueDatInv()
            End If
            MstrNombreCuenta = String.Empty
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Friend ReadOnly Property StrNombreCuenta As String
        Get
            Dim lstrNomCuenta As String
            If HblnEsValido Then
                lstrNomCuenta = MstrNombreCuenta
            Else
                lstrNomCuenta = String.Empty
            End If
            Return lstrNomCuenta
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
Friend Class ClsIdCuentaBancoShr
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsCuentaBanco = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "IdCuentaBanco"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdCuentaBanco"
        HenuTipoValor = EnuTipoValor.enuUShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 2
    End Sub
    Public Overrides Sub SValide()
        HstrMens = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue,
                       BlnEsRequerido, EnuTipoValor)
        If HblnEsValido Then
            If Not BlnLeyendoOrigen Then
                If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                    Dim lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
                    If MobjPadre.FblnExisteLlave(lobjValorLlave) Then
                        HstrMens = "La Cuenta de Banco con el número de identificación ingresado, '" &
                                HobjValorNew.ToString & "', ya existe!"
                        HblnEsValido = False
                    End If
                ElseIf MobjPadre.EnuEstadoActualizacion =
                        EnuEstadoObjetoDef.enuModificando Then
                    HblnEsValido = (HobjValorOriginal = HobjValorNew)
                    If Not HblnEsValido Then
                        HstrMens = "No es permitido cambiar la identidad a objeto alguno!"
                    End If
                End If
            End If
        Else
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                HstrMens = "El valor ingresado, '" & HobjValorNew.ToString & "', no es valido!"
            End If
        End If
        If Not String.IsNullOrEmpty(HstrMens) Then
            SLevanteEveNot("", 0, EnuSeveridadNot.EnuInformacion)
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
Friend Class ClsNombreBancoStr
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsCuentaBanco = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "NombreBanco"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "NombreBanco"
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HshrLongitud = 50
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 3, HshrLongitud,
                HblnEsRequerido)
        HstrMens = String.Empty
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                HblnEsValido = (Not HobjValorNew.ToString.Contains(","))
                If Not HblnEsValido Then
                    HstrMens = "El Nombre del Banco no puede contener comas!"
                    SNotifiqueDatInv()
                End If
            End If
            If HblnEsValido Then
                HobjValorNew = FstrNombreTercero(HobjValorNew)
            End If
        Else
            HstrMens = "El Nombre del Banco debe tener mínimo tres Caracteres!"
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
Friend Class ClsNumeroCuentaStr
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsCuentaBanco = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "NumeroCuenta"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "NumeroCuenta"
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HshrLongitud = 40
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 3, HshrLongitud,
                HblnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuModificando Then
                HblnEsValido = (Not HobjValorNew.ToString.Contains(","))
                If Not HblnEsValido Then
                    HstrMens = "El Número de la Cuenta Bancaria no puede contener comas!"
                    SNotifiqueDatInv()
                End If
            Else
                HobjValorNew = HobjValorNew.ToString.ToUpper
            End If
        Else
            HstrMens = "El Número de la Cuenta Bancaria debe tener mínimo tres Caracteres!"
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
#End Region

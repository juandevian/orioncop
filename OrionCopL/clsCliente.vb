Imports System.Text
Friend Class ClsCliente
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriClientes"
    '
    Private MstrPrediosDelCliente As String() = Nothing
    Private MstrPrediosAgruDelCliente As String() = Nothing
    Private McolPrediosdelCliente As Collection = Nothing
    Private MdtbPrediosDelCliente As DataTable = Nothing
    Private McolItemsProgramaFact As Collection = Nothing
    Private MdtbItemsProgFact As DataTable = Nothing
    Private MdtbFacturas As DataTable = Nothing
    Private MdtbFacturasVivas As DataTable = Nothing
    Private MstrPredioAgruFacturas As String = String.Empty
    Private McolNotasDb As Collection = Nothing
    Private MdtbNotasDb As DataTable = Nothing
    Private MdtbNotasDbExtendida As DataTable = Nothing
    Private MdtbNotasCon As DataTable = Nothing
    Private MdtbNotasAjuste As DataTable = Nothing
    Private MstrPredioAgruNotasDb As String = String.Empty
    Private MblnCargueDtbSoloEstrucura As Boolean = False
    Private MdtbFacturasDeuda As DataTable = Nothing
    Private MdtbAnticipos As DataTable = Nothing
    Private MdtbNotasDevAnt As DataTable = Nothing
    Private MdtbRecibos As DataTable = Nothing
    Private MdtbNotasCr As DataTable = Nothing
    Private MdtbNovedadesDeuda As DataTable = Nothing
    Private MdtbNovedadesAnticipo As DataTable = Nothing
    Private MobjUbicacion As ClsUbicacion = Nothing
    Private MobjEstadoCtaHoy As ClsEstadoCuenta = Nothing
    ' Ultima Deuda calculada
    Private MdecDeudaTotal As Decimal = 0
#End Region

#Region "Constructores"
    ''' <summary>
    ''' Instancia un objeto Cliente.
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
            HcolFiltros.Add(ClsOrionCop.StrFiltroUbicacion)
            lstrCamposSelect = {StrCampoCarpeta, StrCampoCentroUtil,
                   ClsIdClienteDbl.SstrNombreCampoBd}
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
    Public Sub New()
        Dim lstrCamposSelect As String()
        HobjPadre = Nothing
        HcolFiltros.Add(ClsOrionCop.StrFiltroUbicacion)
        lstrCamposSelect = {StrCampoCarpeta, StrCampoCentroUtil,
               ClsIdClienteDbl.SstrNombreCampoBd}
        HblnEsCreable = False
        HblnEsModificable = True
        HblnEsSuprimible = False
        HblnEsAnulable = False
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
            Return EnuIdClasesPanDef.enuCliente
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Cliente"
        End Get
    End Property
    Friend Overrides ReadOnly Property HstrNombreObj As String
        Get
            Return Chr(34) & ObjNombreCompletoStr.ObjValorPro & Chr(34)
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjComentario_ClienteStr As New ClsComentario_ClienteStr(Me)
    Friend ReadOnly Property ObjEsAgenteReteFteBln As New ClsEsAgenteReteFteBln(Me)
    Friend ReadOnly Property ObjEsAutorretenedorBln As New ClsEsAutorretenedorBln(Me)
    Friend ReadOnly Property ObjEsGranContrBln As New ClsEsGranContrBln(Me)
    Friend ReadOnly Property ObjEsRegimenSimpleTBln As New ClsEsRegimenSimpleTBln(Me)
    Friend ReadOnly Property ObjFactPorServicio_CliBln As New ClsFactPorServicio_CliBln(Me)
    Friend ReadOnly Property ObjFechaIngresoDtm As New ClsFechaCreacionDtm(Me)
    Friend ReadOnly Property ObjIdCarpetaClienteShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtilClienteShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdEstadoDeudaByt As New ClsIdEstadoDeudaByt(Me)
    Friend ReadOnly Property ObjIdClienteDbl As New ClsIdClienteDbl(Me)
    Friend ReadOnly Property ObjIdMedioPagoClienteByt As New ClsIdMedioPagoClienteByt(Me)
    Friend ReadOnly Property ObjIdRegimenVentasByt As New ClsIdRegimenVentasByt(Me)
    Friend ReadOnly Property ObjNombreCompletoStr As New ClsNombreCompletoStr(Me)
    Friend ReadOnly Property ObjRecibeDocsPorEmailBln As New ClsRecibeDocsPorEmailBln(Me)
    Friend ReadOnly Property ObjRetieneIcaBln As New ClsRetieneIcaBln(Me)
    Friend ReadOnly Property ObjRetieneIvaBln As New ClsRetieneIvaBln(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjComentario_ClienteStr)
                HcolPropiedades.Add(ObjEsAgenteReteFteBln)
                HcolPropiedades.Add(ObjEsAutorretenedorBln)
                HcolPropiedades.Add(ObjEsGranContrBln)
                HcolPropiedades.Add(ObjEsRegimenSimpleTBln)
                HcolPropiedades.Add(ObjFactPorServicio_CliBln)
                HcolPropiedades.Add(ObjFechaIngresoDtm)
                HcolPropiedades.Add(ObjIdCarpetaClienteShr)
                HcolPropiedades.Add(ObjIdCentroUtilClienteShr)
                HcolPropiedades.Add(ObjIdEstadoDeudaByt)
                HcolPropiedades.Add(ObjIdClienteDbl)
                HcolPropiedades.Add(ObjIdMedioPagoClienteByt)
                HcolPropiedades.Add(ObjIdRegimenVentasByt)
                HcolPropiedades.Add(ObjNombreCompletoStr)
                HcolPropiedades.Add(ObjRecibeDocsPorEmailBln)
                HcolPropiedades.Add(ObjRetieneIcaBln)
                HcolPropiedades.Add(ObjRetieneIvaBln)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras Propiedades"
    Friend ReadOnly Property EnuMedioPagoDefecto As EnuTipoMedioPagoDef
        Get
            Dim lenuMedioPago = ObjIdMedioPagoClienteByt.ObjValorPro
            If lenuMedioPago = EnuTipoMedioPagoDef.None Then
                lenuMedioPago = GobjParametros.ObjIdMedioPagoDefectoByt.ObjValorPro
            End If
            Return lenuMedioPago
        End Get
    End Property
    Friend ReadOnly Property StrResponsFiscal As String
        Get
            Dim lstrResFis = String.Empty
            If ObjEsGranContrBln.ObjValorPro Then
                lstrResFis = "O-13"
            End If
            If ObjEsAutorretenedorBln.ObjValorPro Then
                If Not String.IsNullOrEmpty(lstrResFis) Then
                    lstrResFis &= ";"
                End If
                lstrResFis &= "O-15"
            End If
            If ObjRetieneIvaBln.ObjValorPro Then
                If Not String.IsNullOrEmpty(lstrResFis) Then
                    lstrResFis &= ";"
                End If
                lstrResFis &= "O-23"
            End If
            If ObjEsRegimenSimpleTBln.ObjValorPro Then
                If Not String.IsNullOrEmpty(lstrResFis) Then
                    lstrResFis &= ";"
                End If
                lstrResFis &= "O-47"
            End If
            If String.IsNullOrEmpty(lstrResFis) Then
                lstrResFis = "R-99-PN"
            End If
            Return lstrResFis
        End Get
    End Property
#End Region
#End Region

#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SModifique()
        ObjTerceroCliente.SModifique()
        MyBase.SModifique()
    End Sub
    Protected Overrides Sub SVacie()
        ObjTerceroCliente.SVacie()
        MyBase.SVacie()
        SVacieCompl()
        MdtbFacturas = Nothing
        MdtbFacturasVivas = Nothing
        MstrPrediosAgruDelCliente = Nothing
        MstrPredioAgruFacturas = Nothing
        MstrPredioAgruNotasDb = Nothing
        MstrPrediosDelCliente = Nothing
        McolItemsProgramaFact = Nothing
        McolPrediosdelCliente = Nothing
        MdtbItemsProgFact = Nothing
        MdtbPrediosDelCliente = Nothing
        MdtbFacturasDeuda = Nothing
        MblnCargueDtbSoloEstrucura = False
        MdecDeudaTotal = 0
        MdtbAnticipos = Nothing
        MdtbNotasDb = Nothing
        McolNotasDb = Nothing
        MdtbNotasCon = Nothing
        MdtbNotasDevAnt = Nothing
        MdtbRecibos = Nothing
        MdtbNotasDbExtendida = Nothing
        MdtbNotasCr = Nothing
        MdtbNovedadesDeuda = Nothing
        MdtbNovedadesAnticipo = Nothing
        MdtbNotasAjuste = Nothing
        MobjEstadoCtaHoy = Nothing
    End Sub
    Protected Overrides Sub SInicialiceObj()
        MyBase.SInicialiceObj()
        For Each lobjProp As ClsCBPropiedad In ObjTerceroCliente.ColPropiedades
            If Not lobjProp.BlnEsRequerido Then
                lobjProp.ObjValorPro = ClsPanorama.FobjValorNuloPropiedad(lobjProp)
            End If
        Next
        ObjEsAgenteReteFteBln.ObjValorPro = False
        ObjEsAutorretenedorBln.ObjValorPro = False
        ObjEsGranContrBln.ObjValorPro = False
        ObjEsRegimenSimpleTBln.ObjValorPro = False
        ObjRetieneIcaBln.ObjValorPro = False
        ObjRetieneIvaBln.ObjValorPro = False
        ObjIdCarpetaClienteShr.ObjValorPro = GshrIdCarpeta
        ObjIdCentroUtilClienteShr.ObjValorPro = GshrIdCentroUtil
        ObjIdEstadoDeudaByt.ObjValorPro = EnuEstadoDeudaDef.EnuNormal
        ObjIdMedioPagoClienteByt.ObjValorPro =
                GobjParametros.ObjIdMedioPagoDefectoByt.ObjValorPro
        ObjFactPorServicio_CliBln.ObjValorPro = False
    End Sub
    Protected Overrides Sub SLeaValores(ablnLeyendoOrigen As Boolean)
        MyBase.SLeaValores(ablnLeyendoOrigen)
        SAbraTerceroCliente()
        ObjRecibeDocsPorEmailBln.SValide()
    End Sub
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        Dim lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            GobjPanDat.SInicialiceTransaccion()
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                ObjIdCarpetaClienteShr.ObjValorPro = GshrIdCarpeta
                ObjIdCentroUtilClienteShr.ObjValorPro = GshrIdCentroUtil
                ObjFechaIngresoDtm.ObjValorPro = Date.Now
            End If
            ObjTerceroCliente.SActualice(ablnExigeRequeridos)
            ObjNombreCompletoStr.ObjValorPro = ObjTerceroCliente.FstrNombreCompleto()
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
            If Not lblnNoHayError Then
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            Else
                GobjPanDat.SConfirmeTransaccion()
                GobjPanDat.SControleProcesoObj(False)
            End If
        End Try
    End Sub
    Public Overrides Sub SNormaliceEstado(ablnRefresqueObjeto As Boolean)
        Dim lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            GobjPanDat.SInicialiceTransaccion()
            ObjTerceroCliente.SNormaliceEstado(ablnRefresqueObjeto)
            MyBase.SNormaliceEstado(ablnRefresqueObjeto)
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
            If Not lblnNoHayError Then
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            Else
                GobjPanDat.SConfirmeTransaccion()
                GobjPanDat.SControleProcesoObj(False)
            End If
        End Try
    End Sub
    Protected Overrides Function FblnSuprimio() As Boolean
        Dim lblnSuprimio = FblnEsSuprimible()
        If lblnSuprimio Then
            lblnSuprimio = MyBase.FblnSuprimio()
            If lblnSuprimio Then
                If BlnEsNavegable Then
                    GobjPanorama.SRegistreAccionLogApp(HstrNombreClase, "Suprimir Cliente con Id. " &
                        ObjIdClienteDbl.ToString)
                End If
            End If
        End If
        Return lblnSuprimio
    End Function
    Friend Overrides Function FblnEsSuprimible() As Boolean
        Dim lblnEsSuprimible = FblnPermitidoSuprimir()
        If lblnEsSuprimible Then
            Dim lstrCondicion = " = " & ObjIdClienteDbl.ToString & " AND " &
                    ClsOrionCop.StrFiltroUbicacion
            lblnEsSuprimible = ClsPanorama.FblnEsEliminableReg({SstrNombreTabla},
                    ObjIdClienteDbl.StrNombreCampoBD, lstrCondicion, True, False)
        End If
        Return lblnEsSuprimible
    End Function
    Friend Overrides Function FblnSonValidosDatosOrigen(adtbOrigen As DataTable,
            astrColumnasRelacionadas As String(), ablnReinicie As Boolean,
            ByRef astrMens As String) As Boolean
        Dim lblnEsValido = False, i = 0, lstrColumnaOrigen As String
        Dim lbytIdCar As Byte, lbytIdCenutil As Byte, ldblIdTerCliente As Double
        Dim lobjTer As New ClsTercero(EnuModoInstanciaObjDef.EnuUnico)
        Dim lobjValorLlave As Object()
        For Each ldrwOrigen As DataRow In adtbOrigen.Rows
            i += 1
            lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                    ClsIdCarpetaShr.SstrNombreCampoBd)
            lbytIdCar = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrColumnaOrigen),
                    EnuTipoValor.EnuString)
            lblnEsValido = lbytIdCar = GshrIdCarpeta
            If Not lblnEsValido Then
                astrMens = "La Carpeta del Registro " & i.ToString & " no es la Carpeta actual!"
                Exit For
            End If
            lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                    ClsIdCentroUtilShr.SstrNombreCampoBd)
            lbytIdCenutil = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrColumnaOrigen),
                    EnuTipoValor.EnuString)
            lblnEsValido = lbytIdCenutil = GshrIdCentroUtil
            If Not lblnEsValido Then
                astrMens = "La Copropiedad del Registro " & i.ToString & " no es el actual!"
                Exit For
            End If
            lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                    ClsIdClienteDbl.SstrNombreCampoBd)
            ldblIdTerCliente = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrColumnaOrigen),
                    EnuTipoValor.EnuDouble)
            lblnEsValido = ClsPanorama.FblnEsValidoNumero(ldblIdTerCliente, 1, Double.MaxValue,
                    True, EnuTipoValor.EnuDouble)
            If Not lblnEsValido Then
                astrMens = "La Identificación del Cliente en el Registro " & i.ToString &
                        " no es válido!"
                Exit For
            End If
            lobjValorLlave = {ldblIdTerCliente}
            lobjTer.SAbra(lobjValorLlave)
            lblnEsValido = lobjTer.BlnExiste
            If Not lblnEsValido Then
                astrMens = "El Tercero correspondiente al Cliente del registro " & i.ToString &
                        " no existe!"
                Exit For
            End If
        Next
        Return lblnEsValido
    End Function
#End Region

#Region "Procedimientos del objeto"
    Friend Sub SModifiqueParaEstado()
        HblnEsAnulable = False
        HblnEsCreable = False
        HblnEsSuprimible = False
        HblnEsModificable = False
    End Sub
    Friend Function FstrNombrePaisDir() As String
        Dim lstrNomPais As String
        If IsNothing(MobjUbicacion) Then
            MobjUbicacion = New Ubicacion.ClsUbicacion
        End If
        lstrNomPais = MobjUbicacion.StrNombrePais(ObjPaisDirStr.ObjValorPro)
        Return FstrComillas(lstrNomPais)
    End Function
    ''' <summary>
    ''' Devueleve el código del departamento si el pais es Colombia de lo contrario devuelve
    ''' el string null
    ''' </summary>
    ''' <returns></returns>
    Friend Function FstrIdSubdivisionPais() As String
        Dim lstrSubPais = String.Empty
        If ObjPaisDirStr.ObjValorPro = "CO" Then
            lstrSubPais = Format(ObjDepartamentoDirByt.ObjValorPro, "0#")
        End If
        lstrSubPais = FstrComillas(lstrSubPais)
        Return lstrSubPais
    End Function
    Friend Function FstrNombreSubPais() As String
        Dim lstrNomSubPais = String.Empty
        If ObjPaisDirStr.ObjValorPro = "CO" Then
            lstrNomSubPais = MobjUbicacion.StrNombreDpto(ObjPaisDirStr.ObjValorPro,
                    ObjDepartamentoDirByt.ObjValorPro)
        End If
        Return FstrComillas(lstrNomSubPais)
    End Function
    ''' <summary>
    ''' Devueleve el código de la ciudad si el pais es Colombia de lo contrario devuelve
    ''' el string null
    ''' </summary>
    ''' <returns></returns>
    Friend Function FstrIdCiudad() As String
        Dim lstrIdCiu = String.Empty
        If ObjPaisDirStr.ObjValorPro = "CO" Then
            Dim lstrIdDpto = Format(ObjDepartamentoDirByt.ObjValorPro, "0#")
            lstrIdCiu = Format(ObjCiudadDirShr.ObjValorPro, "00#")
            lstrIdCiu = lstrIdDpto & lstrIdCiu
        End If
        lstrIdCiu = FstrComillas(lstrIdCiu)
        Return lstrIdCiu
    End Function
    Friend Function FstrNombreCiudad() As String
        Dim lstrNomCiudad = String.Empty
        If ObjPaisDirStr.ObjValorPro = "CO" Then
            lstrNomCiudad = MobjUbicacion.StrNombreCiudad(ObjPaisDirStr.ObjValorPro,
                    ObjDepartamentoDirByt.ObjValorPro, ObjCiudadDirShr.ObjValorPro)
        End If
        Return FstrComillas(lstrNomCiudad)
    End Function
    Friend Function FstrTelCliente() As String
        Dim lstrTel As String = ObjCelularStr.ToString
        If String.IsNullOrEmpty(lstrTel) OrElse lstrTel = "0" Then
            lstrTel = ObjCelular2Str.ToString
        End If
        If String.IsNullOrEmpty(lstrTel) OrElse lstrTel = "0" Then
            lstrTel = ObjTelefonoUnoStr.ToString
        End If
        If String.IsNullOrEmpty(lstrTel) OrElse lstrTel = "0" Then
            lstrTel = ObjTelefonoDosStr.ToString
        End If
        If lstrTel = "0" Then
            lstrTel = String.Empty
        End If
        lstrTel = FstrComillas(lstrTel)
        Return lstrTel
    End Function
    Friend Function FstrRegimenIva() As String
        Dim lstrRegIva As String
        If ObjIdRegimenVentasByt.ObjValorPro = EnuRegimenVentasDef.enuResponsable Then
            lstrRegIva = "48"
        Else
            lstrRegIva = "49"
        End If
        Return FstrComillas(lstrRegIva)
    End Function
#End Region

#Region "Anticipos"
    Friend ReadOnly Property DtbAnticipos(astrIdPredioAgr As String) As DataTable
        Get
            SCargueDtbAnticipos(astrIdPredioAgr, False, False)
            SComplementeDtbAnticipos()
            Return MdtbAnticipos
        End Get
    End Property
    Friend ReadOnly Property DtbAnticiposVivos(astrIdPredioAgr As String) As DataTable
        Get
            SCargueDtbAnticipos(astrIdPredioAgr, False, True)
            SComplementeDtbAnticipos()
            Return MdtbAnticipos
        End Get
    End Property
    Friend Function FobjNuevoAnticipo(adtmFecha As Date, astrIdPredioAgrupador As String) As ClsAnticipo
        SCargueDtbAnticipos(astrIdPredioAgrupador, True, False)
        Dim lblnModificoPermisos = False
        Dim ldrwNuevoAnticipo As DataRow = MdtbAnticipos.NewRow
        Dim lobjNuevoAnticipo As New ClsAnticipo(Me, ldrwNuevoAnticipo)
        With lobjNuevoAnticipo
            If Not CType(.EnuPermisosObj And EnuPermisosDef.enuCrear, Boolean) Then
                .EnuPermisosObj += EnuPermisosDef.enuCrear
                lblnModificoPermisos = True
            End If
            .SCreeObj(Nothing)
            .ObjAnuladoBln.ObjValorPro = False
            .ObjDebitos_AntDec.ObjValorPro = 0
            .ObjFechaAnticipoDtm.ObjValorPro = adtmFecha
            .ObjFechaCreacionDtm.ObjValorPro = Date.Now
            .ObjIdCarpeta_AntShr.ObjValorPro = GshrIdCarpeta
            .ObjIdCentroUtil_AntShr.ObjValorPro = GshrIdCentroUtil
            .ObjIdCliente_AntDbl.ObjValorPro = ObjIdClienteDbl.ObjValorPro
            .ObjIdPredioAgrupador_AntStr.ObjValorPro = astrIdPredioAgrupador
            If lblnModificoPermisos Then
                .EnuPermisosObj -= EnuPermisosDef.enuCrear
            End If
        End With
        Return lobjNuevoAnticipo
    End Function
    Private Sub SCargueDtbAnticipos(astrIdPredioAgr As String, ablnSoloEstructura As Boolean,
             ablnSoloConSaldo As Boolean)
        Static lblnSoloEstructura As Boolean = False
        Static lstrIdPredioAgr As String = String.Empty
        If IsNothing(MdtbAnticipos) OrElse (lblnSoloEstructura <> ablnSoloEstructura) OrElse
                lstrIdPredioAgr <> astrIdPredioAgr Then
            lstrIdPredioAgr = astrIdPredioAgr
            lblnSoloEstructura = ablnSoloEstructura
            Dim lstrAnu = "IIF(" & ClsAnuladoBln.SstrNombreCampoBd & " = 0, 'No', 'Si') AS Anu"
            Dim lstrIndice = {{StrCampoCarpeta, "ASC"},
                              {StrCampoCentroUtil, "ASC"},
                              {ClsIdPredioAgrupador_AntStr.SstrNombreCampoBd, "ASC"},
                              {ClsIdAnticipo_NovEnt.SstrNombreCampoBd, "ASC"}}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsIdCliente_AntDbl.SstrNombreCampoBd & " = "
            If ablnSoloEstructura Then
                lstrFiltro &= "0 AND " & ClsIdPredioAgrupador_AntStr.SstrNombreCampoBd & " = ''"
            Else
                lstrFiltro &= ObjIdClienteDbl.ObjValorPro & " AND " &
                        ClsIdPredioAgrupador_AntStr.SstrNombreCampoBd & " = '" & astrIdPredioAgr & "'"
            End If
            If ablnSoloConSaldo Then
                lstrFiltro &= " AND " & ClsCreditos_AntDec.SstrNombreCampoBd & " - " & ClsDebitos_AntDec.SstrNombreCampoBd &
                        " > 0 "
            End If
            Dim lstrCamposSelect() = {"*", "'' As Servicio", "'' AS NroDocOriAnt", "'' AS DocOrigen",
                    "0 as Saldo", lstrAnu}
            MdtbAnticipos = ClsPanorama.FdtbDataTable(ClsAnticipo.SstrNombreTabla, lstrCamposSelect,
                    lstrIndice, lstrFiltro)
        End If
    End Sub
    Private Sub SComplementeDtbAnticipos()
        If Not IsNothing(MdtbAnticipos) Then
            Dim ldrwAnticipos As DataRow() = MdtbAnticipos.Select()
            Dim lstrPrefDocOriAnt As String, lentIdDocOriAnt As Integer, lstrNroDocOriAnt As String
            Dim ldecCreditosAnt As Decimal, ldecDebitos As Decimal, ldecSaldo As Decimal
            Dim lstrServicios As String, lstrServicio = String.Empty
            Dim lenuDocOrigen As EnuTipoDocOri
            For Each ldrwAnt As DataRow In ldrwAnticipos
                lstrServicios = ClsPanorama.FobjValorCampo(ldrwAnt(
                        ClsServicios_AntStr.SstrNombreCampoBd), EnuTipoValor.enuString)
                lstrPrefDocOriAnt = ClsPanorama.FobjValorCampo(
                        ldrwAnt(ClsPrefijoDocOrigen_AntStr.SstrNombreCampoBd),
                        EnuTipoValor.enuString)
                lentIdDocOriAnt = ClsPanorama.FobjValorCampo(
                        ldrwAnt(ClsIdDocOrigen_AntEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
                lstrNroDocOriAnt = ClsPanorama.FstrNumeroDcto(lstrPrefDocOriAnt, lentIdDocOriAnt)
                ldecCreditosAnt = ClsPanorama.FobjValorCampo(ldrwAnt(
                        ClsCreditos_AntDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
                ldecDebitos = ClsPanorama.FobjValorCampo(ldrwAnt(
                        ClsDebitos_AntDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
                ldecSaldo = ldecCreditosAnt - ldecDebitos
                lenuDocOrigen = ClsPanorama.FobjValorCampo(ldrwAnt(
                        ClsIdTipoDocOrigen_AntByt.SstrNombreCampoBd), EnuTipoValor.enuByte)
                If lenuDocOrigen = EnuTipoDocOri.EnuReciboCaja Then
                    ldrwAnt("DocOrigen") = "REC"
                Else
                    ldrwAnt("DocOrigen") = "NAJ"
                End If
                ldrwAnt("NroDocOriAnt") = lstrNroDocOriAnt
                ldrwAnt("Saldo") = ldecSaldo
                If lstrServicios = "A" Then
                    lstrServicio = "TODOS"
                ElseIf lstrServicios = "0" Then
                    lstrServicio = "Cuotas de Administración"
                ElseIf lstrServicios.Contains(",") Then
                    lstrServicio = "Varios Servicios"
                Else
                    If IsNumeric(lstrServicios) Then
                        lstrServicio = GobjParametros.FstrNombreServicio(lstrServicios)
                    End If
                End If
                ldrwAnt("Servicio") = lstrServicio
            Next
        End If
    End Sub
#End Region

#Region "Causación Mora"
    ' Causa mora a predios agrupadores en procesos FM
    Friend Function SCauseMora(adtmFechaCausacion As Date, ByRef astrMens As String) As Decimal
        Dim ldecMoraCausada = 0D, lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            GobjPanDat.SInicialiceTransaccion()
            If GobjParametros.FdblTasaMoraFecha(adtmFechaCausacion) > 0 Then
                Dim lstrPrediosAgruFact() As String = FstrPrediosAgrupEnFras()
                If lstrPrediosAgruFact.Length > 0 Then
                    ldecMoraCausada = SCauseMoraPreAgr(lstrPrediosAgruFact,
                            adtmFechaCausacion, astrMens)
                End If
            End If
            lblnNoHayError = True
        Catch ex As ErrorInesperadoPanDatException
            Throw
        Catch ex As ErrorInesperadoPanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As PanLException
            Throw
        Catch ex As ArgumentException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                GobjPanDat.SConfirmeTransaccion()
                GobjPanDat.SControleProcesoObj(False)
                SRefresqueObj()
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
        Return ldecMoraCausada
    End Function
    Friend Function SCauseMora(astrIdPrediosAgru As String(), adtmFechaCausacion As Date,
            ByRef astrMens As String) As Decimal
        Dim ldecMoraCausada = 0D, lblnNoHayError = False
        Try
            If GobjParametros.FdblTasaMoraFecha(adtmFechaCausacion) > 0 Then
                GobjPanDat.SControleProcesoObj(True)
                GobjPanDat.SInicialiceTransaccion()
                Dim ldtmFechaUltCau As Date = GobjParametros.ObjFechaUltCausacionGralDtm.ObjValorPro
                ldecMoraCausada = SCauseMoraPreAgr(astrIdPrediosAgru, adtmFechaCausacion,
                        astrMens)
            End If
            lblnNoHayError = True
        Catch ex As ErrorInesperadoPanDatException
            Throw
        Catch ex As ErrorInesperadoPanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As PanLException
            Throw
        Catch ex As ArgumentException
            Throw
        Catch ex As Exception
            Throw
        Finally
            If lblnNoHayError Then
                GobjPanDat.SConfirmeTransaccion()
                GobjPanDat.SControleProcesoObj(False)
                SRefresqueObj()
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
        Return ldecMoraCausada
    End Function
    ' Causa mora a predios agrupadores del cliente en procesos FM
    Private Function SCauseMoraPreAgr(astrIdPrediosAgru() As String, adtmFechaCausacion As Date,
            ByRef astrMens As String) As Decimal
        Dim lstrIdPreAgr As String, lstrIdPreAgrActual As String = "***"
        Dim lobjNuevaNotaDb As ClsNotaDb = Nothing
        Dim ldecIntMoraFact As Decimal, lblnNoPudoCausar = False
        Dim ldecIntMoraCausados As Decimal
        Dim lstrServicios As String() = {}
        Dim lcolFrasVivasIva As New Collection
        Dim lcolFrasVivas = FcolFacturas(astrIdPrediosAgru, lstrServicios, True)
        Dim ldtmFechaNota As Date
        Dim ldtmFechaInicioPer = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
        For Each lobjFac As ClsFactura In lcolFrasVivas
            If lobjFac.FdecIvaCapital > 0 Then
                lcolFrasVivasIva.Add(lobjFac)
                Continue For
            End If
            If lobjFac.ObjFechaVencimientoDtm.ObjValorPro < adtmFechaCausacion Then
                lstrIdPreAgr = lobjFac.ObjIdPredioAgrupador_FacStr.ObjValorPro
                ' Verifica la cuenta no está suspendida o perdída
                If FblnPuedeCausarMora(lstrIdPreAgr) Then
                    If GblnCausandoFM Then
                        Dim lenuModoCM As EnuModoCausaMora = lobjFac.FenuModoCausaMora
                        ldtmFechaNota = If(lenuModoCM = EnuModoCausaMora.EnuUltimoDia,
                                ldtmFechaInicioPer.AddDays(-1), adtmFechaCausacion)
                    Else
                        ldtmFechaNota = adtmFechaCausacion
                    End If
                    lobjFac.EnuPermisosObj += EnuPermisosDef.enuModificar
                    If lstrIdPreAgrActual = "***" Then
                        lobjNuevaNotaDb = FobjNuevaNotaDb(lstrIdPreAgr, EnuOrigenNotaDb.EnuAplicacion)
                        If Not lobjNuevaNotaDb.EnuPermisosObj And EnuPermisosDef.enuCrear Then
                            lobjNuevaNotaDb.EnuPermisosObj += EnuPermisosDef.enuCrear
                        End If
                        lstrIdPreAgrActual = lstrIdPreAgr
                    End If
                    If lstrIdPreAgr <> lstrIdPreAgrActual Then
                        If lobjNuevaNotaDb.ColItemsNotaDb.Count > 0 Then
                            lobjNuevaNotaDb.ObjFecha_NotaDbDtm.ObjValorPro = ldtmFechaNota
                            lobjNuevaNotaDb.SActualice(True)
                        End If
                        lstrIdPreAgrActual = lstrIdPreAgr
                        ldecIntMoraCausados = 0
                        lobjNuevaNotaDb = FobjNuevaNotaDb(lstrIdPreAgr, EnuOrigenNotaDb.EnuAplicacion)
                        If Not lobjNuevaNotaDb.EnuPermisosObj And EnuPermisosDef.enuCrear Then
                            lobjNuevaNotaDb.EnuPermisosObj += EnuPermisosDef.enuCrear
                        End If
                    End If
                    ldecIntMoraFact = lobjFac.SCauseMora(adtmFechaCausacion)
                    If lobjFac.FblbModificoItem Then
                        If ldecIntMoraFact > 0 Then
                            ldecIntMoraCausados += ldecIntMoraFact
                            SGenereItemsNDb(lobjNuevaNotaDb, lobjFac,
                                GobjParametros.FdblTasaMoraFecha(adtmFechaCausacion))
                        End If
                        lobjFac.SActualice(True)
                    End If
                Else
                    lblnNoPudoCausar = True
                End If
            End If
        Next
        If lobjNuevaNotaDb IsNot Nothing AndAlso lobjNuevaNotaDb.ColItemsNotaDb.Count > 0 Then
            lobjNuevaNotaDb.ObjFecha_NotaDbDtm.ObjValorPro = ldtmFechaNota
            lobjNuevaNotaDb.SActualice(True)
        End If
        If lcolFrasVivasIva.Count > 0 Then
            ldecIntMoraFact = FdecCausoMoraFactsIva(lcolFrasVivasIva, adtmFechaCausacion,
                    astrMens)
            ldecIntMoraCausados += ldecIntMoraFact
        End If
        If lblnNoPudoCausar Then
            astrMens = "No se Causo intereses a Deudas Suspendidas y/o Perdidas!"
        End If
        Return ldecIntMoraCausados
    End Function
    Private Function FdecCausoMoraFactsIva(acolFacturasIva As Collection, adtmFechaCausacion As Date,
            ByRef astrMens As String) As Decimal
        Dim lstrIdPreAgr As String, ldecIntMoraFact As Decimal, lblnNoPudoCausar = False
        Dim ldtmFechaNota As Date, ldecIntMoraCausados As Decimal
        Dim ldtmFechaInicioPer = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
        Dim lobjNuevaNotaDb As ClsNotaDb
        For Each lobjFac As ClsFactura In acolFacturasIva
            lstrIdPreAgr = lobjFac.ObjIdPredioAgrupador_FacStr.ObjValorPro
            ' Verifica la cuenta no está suspendida o perdída
            If FblnPuedeCausarMora(lstrIdPreAgr) Then
                If GblnCausandoFM Then
                    Dim lenuModoCM As EnuModoCausaMora = lobjFac.FenuModoCausaMora
                    ldtmFechaNota = If(lenuModoCM = EnuModoCausaMora.EnuUltimoDia,
                                ldtmFechaInicioPer.AddDays(-1), adtmFechaCausacion)
                Else
                    ldtmFechaNota = adtmFechaCausacion
                End If
                lobjFac.EnuPermisosObj += EnuPermisosDef.enuModificar
                lobjNuevaNotaDb = FobjNuevaNotaDb(lstrIdPreAgr, EnuOrigenNotaDb.EnuAplicacion)
                If Not lobjNuevaNotaDb.EnuPermisosObj And EnuPermisosDef.enuCrear Then
                    lobjNuevaNotaDb.EnuPermisosObj += EnuPermisosDef.enuCrear
                End If
                ldecIntMoraFact = lobjFac.SCauseMora(adtmFechaCausacion)
                If lobjFac.FblbModificoItem Then
                    If ldecIntMoraFact > 0 Then
                        ldecIntMoraCausados += ldecIntMoraFact
                        SGenereItemsNDb(lobjNuevaNotaDb, lobjFac,
                            GobjParametros.FdblTasaMoraFecha(adtmFechaCausacion))
                    End If
                    lobjFac.SActualice(True)
                End If
                If lobjNuevaNotaDb IsNot Nothing AndAlso lobjNuevaNotaDb.ColItemsNotaDb.Count > 0 Then
                    lobjNuevaNotaDb.ObjFecha_NotaDbDtm.ObjValorPro = ldtmFechaNota
                    lobjNuevaNotaDb.SActualice(True)
                    lobjNuevaNotaDb = Nothing
                End If
            Else
                lblnNoPudoCausar = True
            End If
        Next
        If lblnNoPudoCausar Then
            astrMens = "No se Causo intereses a Deudas Suspendidas y/o Perdidas!"
        End If
        Return ldecIntMoraCausados
    End Function
    ''' <summary>
    ''' Indica si la deuda del cliente y del predio agrupador no esta suspendida o perdida
    ''' </summary>
    ''' <param name="astrIdPredio"></param>
    ''' <returns>True si estado de deuda no es ni suspendida ni vencida y false si lo contrario</returns>
    Friend Function FblnPuedeCausarMora(astrIdPredio As String) As Boolean
        Dim lblnPuede As Boolean
        If String.IsNullOrEmpty(astrIdPredio) Then
            lblnPuede = ObjIdEstadoDeudaByt.ObjValorPro < EnuEstadoDeudaDef.EnuPerdida
        Else
            Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
            lobjPredio.SAbra({GshrIdCarpeta, GshrIdCentroUtil, astrIdPredio})
            lblnPuede = lobjPredio.ObjIdEstadoDeuda_PredioByt.ObjValorPro <
                    EnuEstadoDeudaDef.EnuPerdida
        End If
        Return lblnPuede
    End Function
    Private Shared Sub SGenereItemsNDb(aobjNuevaNotaDb As ClsNotaDb, aobjFactura As ClsFactura,
            adblTasaMora As Double)
        For i = 0 To aobjFactura.ColItemsFactura.Count - 1
            If aobjFactura.StcIntMora_DecVlrMora(i) > 0 Then
                aobjNuevaNotaDb.SAdicioneItemNotaDb(aobjFactura.ObjIdFacturaEnt.ObjValorPro,
                        aobjFactura.ObjPrefijo_FactStr.ObjValorPro, aobjFactura.StcIntMoraFactura(i),
                        adblTasaMora)
            End If
        Next
    End Sub
    Private Function FstrPrediosAgrupEnFras() As String()
        Dim lstrIdPrediosAgrup As String() = Array.Empty(Of String)
        Dim lstrCamposSelect() = {"DISTINCT " & ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd}
        Dim lstrIndice(,) = {{ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdCliente_FactDbl.SstrNombreCampoBd & " = " & ObjIdClienteDbl.ObjValorPro & " AND " &
                ClsAnuladoBln.SstrNombreCampoBd & " = FALSE AND " &
                ClsPrefijo_FactStr.SstrNombreCampoBd & " <> '" & GCSTRPREFPREFACTURA & "' AND " &
                ClsDebitos_FactDec.SstrNombreCampoBd & " > " & ClsCreditos_FactDec.SstrNombreCampoBd
        Dim ldtbPrediosAgrup = ClsPanorama.FdtbDataTable(ClsFactura.SstrNombreTabla, lstrCamposSelect,
                                                         lstrIndice, lstrFiltro)
        Dim ldrwPrediosAgru() As DataRow = ldtbPrediosAgrup.Select()
        If ldrwPrediosAgru.Count > 0 Then
            ReDim lstrIdPrediosAgrup(ldrwPrediosAgru.Count - 1)
        End If
        For i = 0 To ldrwPrediosAgru.Count - 1
            lstrIdPrediosAgrup(i) = ClsPanorama.FobjValorCampo(ldrwPrediosAgru(i)(0), EnuTipoValor.enuString)
        Next
        Return lstrIdPrediosAgrup
    End Function
#End Region

#Region "Deuda del Cliente"
    ' Usadas en la creacion de recibos de caja
    ''' <summary>
    ''' Devuelve el total de la deuda de capital para los predios agrupadores y 
    ''' el servicio pasado en los argumentos, mas el total de la deuda por mora 
    ''' para todos los servicios de los predios agrupadores pasados en el argumento
    ''' </summary>
    ''' <param name="astrIdPrediosAgru">Array de los predios agrupadores</param>
    ''' <param name="astrServicios">Array de los servicios donde "0" indica que 
    ''' es Cuota de Administración</param>''' 
    ''' <returns></returns>
    Friend Function FdecDeuda(astrIdPrediosAgru As String(), astrServicios As String()) As Decimal
        MdecDeudaTotal = 0
        Dim ldecDeudaCap = FdecDeudaCapital(astrIdPrediosAgru, astrServicios)
        Dim ldecDeudaMora = FdecDeudaMora(astrIdPrediosAgru)
        MdecDeudaTotal = ldecDeudaCap + ldecDeudaMora
        Return MdecDeudaTotal
    End Function
    ''' <summary>
    ''' Devuelve el total de la deuda de capital para los predios agrupadores y el 
    ''' servicio pasado en el argumento
    ''' </summary>
    ''' <returns></returns>
    Private Function FdecDeudaCapital(astrIdPrediosAgr As String(),
            astrServicios As String()) As Decimal
        Dim lstbSql As New StringBuilder
        Dim ldecDeudaCap As Decimal = 0
        Dim lblnCalcular = MdecDeudaTotal = 0
        If lblnCalcular Then
            Dim lblnultimo = False, i = 0
            Dim lstrFiltroPre = FstrFiltroPredios(astrIdPrediosAgr)
            Dim lstrFilSer = FstrFiltroServi(astrServicios)
            With lstbSql _
                .Append("SELECT ") _
                .Append("SUM(IF(").Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = ") _
                .Append(EnuTipoNov.EnuDbCap).Append(" OR ") _
                .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = ") _
                .Append(EnuTipoNov.EnuDbIva).Append(" OR ") _
                .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = ") _
                .Append(EnuTipoNov.EnuRCrPagoCap).Append(" OR ") _
                .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = ") _
                .Append(EnuTipoNov.EnuRCrAnApCap).Append(" OR ") _
                .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = ") _
                .Append(EnuTipoNov.EnuRCrDctoCap).Append(" OR ") _
                .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd) _
                .Append(" = ").Append(EnuTipoNov.EnuRCrIvaGas).Append(" OR ") _
                .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" BETWEEN ") _
                .Append(EnuTipoNov.EnuRCrRetFte).Append(" AND ") _
                .Append(EnuTipoNov.EnuRCrRetCre).Append(", Valor, 0) - IF(") _
                .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = ") _
                .Append(EnuTipoNov.EnuCrPagoCap).Append(" OR ") _
                .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = ") _
                .Append(EnuTipoNov.EnuCrAnApCap).Append(" OR ") _
                .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = ") _
                .Append(EnuTipoNov.EnuCrDctoCap).Append(" OR ") _
                .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" BETWEEN ") _
                .Append(EnuTipoNov.EnuCrRetFte).Append(" And ") _
                .Append(EnuTipoNov.EnuCrRetCre).Append(" OR ") _
                .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = ") _
                .Append(EnuTipoNov.EnuRDbCap).Append(" OR ") _
                .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = ") _
                .Append(EnuTipoNov.EnuRDbIva).Append(" OR ") _
                .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = ") _
                .Append(EnuTipoNov.EnuCrIvaGas) _
                .Append(", Valor, 0)) As Saldo FROM ").Append(ClsNovedad.SstrNombreTabla) _
                .Append(" WHERE ") _
                .Append(StrCampoCarpeta).Append(" = ").Append(GshrIdCarpeta) _
                .Append(" AND ").Append(StrCampoCentroUtil).Append(" = ") _
                .Append(GshrIdCentroUtil).Append(" AND ").Append(ClsIdClienteDbl.SstrNombreCampoBd) _
                .Append(" = ").Append(ObjIdClienteDbl.ObjValorPro).Append(" AND ")
                .Append(lstrFiltroPre)
                If Not String.IsNullOrEmpty(lstrFilSer) Then
                    .Append(" AND ").Append(lstrFilSer)
                End If
            End With
            Dim ldtbDeuda = ClsPanorama.FdtbDataTable(lstbSql.ToString)
            ldecDeudaCap = ClsPanorama.FobjValorCampo(ldtbDeuda.Rows(0)("Saldo"),
                    EnuTipoValor.enuDecimal)
        End If
        Return ldecDeudaCap
    End Function
    ''' <summary>
    ''' Indica si el cliente tiene facturas en mora
    ''' </summary>
    ''' <returns></returns>
    Friend Function FblnTieneDeuda(ablnEnMora As Boolean) As Boolean
        Dim lblnTiene As Boolean
        Dim lstrFecha = "'" & ClsPanoramaDat.FstrFechaNormalizada(Date.Today) & "'"
        Dim lstrTabla = ClsFactura.SstrNombreTabla
        Dim lstrCamSelect = {"SUM(" & ClsDebitos_FactDec.SstrNombreCampoBd & " - " &
                ClsCreditos_FactDec.SstrNombreCampoBd & ") AS Saldo"}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdCliente_FactDbl.SstrNombreCampoBd & " = " & ObjIdTerceroDbl.ToString &
                " AND "
        If ablnEnMora Then
            lstrFiltro &= ClsFechaVencimientoDtm.SstrNombreCampoBd & " < " & lstrFecha
        Else
            lstrFiltro &= ClsFechaFacturaDtm.SstrNombreCampoBd & " < " & lstrFecha
        End If
        Dim lstrOrden As String(,) = {{"", ""}}
        Dim ldtbRes = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamSelect, lstrOrden, lstrFiltro)
        Dim ldecDeuda As Decimal = ClsPanorama.FobjValorCampo(ldtbRes.Rows(0)(0),
                EnuTipoValor.enuDecimal)
        lblnTiene = ldecDeuda > 0
        Return lblnTiene
    End Function

    ''' <summary>
    ''' Devuelve el total de intereses de mora debidos por los predios agrupadores pasados
    ''' en el argumento .
    ''' </summary>
    ''' <param name="astrIdPrediosAgr">Array con los predios agrupadores</param>
    ''' <returns></returns>
    Friend Function FdecDeudaMora(astrIdPrediosAgr As String()) As Decimal
        Dim ldecDeudaMora As Decimal = 0
        Dim lstbSql As New StringBuilder
        With lstbSql
            .Append("SELECT SUM(IF(").Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = ")
            .Append(EnuTipoNov.EnuDbInt).Append(" OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = ")
            .Append(EnuTipoNov.EnuDbIvaInt).Append(" OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = ")
            .Append(EnuTipoNov.EnuRCrPagoInt).Append(" OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = ")
            .Append(EnuTipoNov.EnuRCrAnApInt).Append(" OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = ")
            .Append(EnuTipoNov.EnuRCrDctoInt)
            .Append(", Valor, 0) - IF(")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = ")
            .Append(EnuTipoNov.EnuCrPagoInt).Append(" OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = ")
            .Append(EnuTipoNov.EnuRDbIvaInt).Append(" OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = ")
            .Append(EnuTipoNov.EnuCrAnApInt).Append(" OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = ")
            .Append(EnuTipoNov.EnuCrDctoInt).Append(" OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = ")
            .Append(EnuTipoNov.EnuRDbInt).Append(", Valor, 0)) AS Saldo FROM ")
            .Append(ClsNovedad.SstrNombreTabla).Append(" WHERE ")
            .Append(ClsIdCarpetaShr.SstrNombreCampoBd).Append(" = ")
            .Append(GshrIdCarpeta).Append(" AND ")
            .Append(ClsIdCentroUtilShr.SstrNombreCampoBd).Append(" = ").Append(GshrIdCentroUtil)
            .Append(" AND ").Append(ClsIdTercero_NovDbl.SstrNombreCampoBd).Append(" = ")
            .Append(ObjIdClienteDbl.ToString)
        End With
        Dim lstrFiltroPreAgr = FstrFiltroPredios(astrIdPrediosAgr)
        Dim lstrSql = lstbSql.ToString & " AND " & lstrFiltroPreAgr
        Dim ldtbDeudaMora = ClsPanorama.FdtbDataTable(lstrSql)
        ldecDeudaMora = ClsPanorama.FobjValorCampo(ldtbDeudaMora.Rows(0)("Saldo"),
            EnuTipoValor.enuDecimal)
        Return ldecDeudaMora
    End Function

    ''' <summary>
    ''' Devuelve un Array con los predios agrupadores del cliente que tienen Anticipos por aplicar.
    ''' </summary>
    Friend Function FdrwPrediosAgruAnt() As DataRow()
        Static ldblIdCliente As Double = 0
        Static ldrwPrediosAgru As DataRow() = Nothing
        If IsNothing(ldrwPrediosAgru) OrElse ldblIdCliente <> ObjIdClienteDbl.ObjValorPro Then
            ldblIdCliente = ObjIdClienteDbl.ObjValorPro
            Dim lstrNombreTabla = ClsAnticipo.SstrNombreTabla
            Dim lstrCamposSelect As String() = {"DISTINCT " & ClsIdPredioAgrupador_AntStr.SstrNombreCampoBd}
            Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsIdCliente_AntDbl.SstrNombreCampoBd & " = " & ObjIdClienteDbl.ObjValorPro &
                    " AND " & ClsCreditos_AntDec.SstrNombreCampoBd & " > " &
                    ClsDebitos_AntDec.SstrNombreCampoBd
            Dim lstrIndice(,) = {{ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd, "ASC"}}
            GobjPanDat.SControleProcesoObj(True)
            Dim ldtbPrediosAgru As DataTable = ClsPanorama.FdtbDataTable(lstrNombreTabla, lstrCamposSelect,
                    lstrIndice, lstrFiltro)
            GobjPanDat.SControleProcesoObj(False)
            ldrwPrediosAgru = ldtbPrediosAgru.Select()
        End If
        Return ldrwPrediosAgru
    End Function
    Friend Function FdecSaldoDeudaFecha(astrIdPredioAgr As String, adtmFechaSaldo As Date) As Decimal
        Dim ldecSaldoDeuda As Decimal
        Dim ldecTotalDb = FdecTotalDebitos(astrIdPredioAgr, adtmFechaSaldo)
        Dim ldecTotalCr = FdecTotalCreditos(astrIdPredioAgr, adtmFechaSaldo)
        ldecSaldoDeuda = ldecTotalDb - ldecTotalCr
        Return ldecSaldoDeuda
    End Function
    Private Function FdecTotalDebitos(astrIdPredioAgr As String, adtmFechaTotal As Date) As Decimal 'Ok Iva Mora
        Dim lstrFecha = "'" & ClsPanoramaDat.FstrFechaNormalizada(adtmFechaTotal) & "'"
        Dim lstrCamposSelect = {"SUM(" & ClsValor_NovDec.SstrNombreCampoBd & ")"}
        Dim lstbFiltro As New StringBuilder
        With lstbFiltro
            .Clear().Append(ClsOrionCop.StrFiltroUbicacion).Append(" AND ")
            .Append(ClsIdTercero_NovDbl.SstrNombreCampoBd).Append(" = ")
            .Append(ObjIdClienteDbl.ObjValorPro).Append(" AND ")
            .Append(ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd).Append(" = '")
            .Append(astrIdPredioAgr).Append("' AND (")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = ")
            .Append(EnuTipoNov.EnuDbIvaInt).Append(" OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" BETWEEN ")
            .Append(EnuTipoNov.EnuDbCap).Append(" AND ").Append(EnuTipoNov.EnuDbInt).Append(" OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" BETWEEN ")
            .Append(EnuTipoNov.EnuRCrPagoCap).Append(" AND ").Append(EnuTipoNov.EnuRCrRetCre)
            .Append(") AND ").Append(ClsFechaNovedadDtm.SstrNombreCampoBd).Append(" <= ")
            .Append(lstrFecha)
        End With
        Dim lstrFiltro = lstbFiltro.ToString
        Dim ldtbTotalDeb = ClsPanorama.FdtbDataTable(ClsNovedad.SstrNombreTabla, lstrCamposSelect, {{}},
                    lstrFiltro)
        Dim ldrwTotalDb As DataRow = ldtbTotalDeb.Select()(0)
        Dim ldecTotalDb As Decimal = ClsPanorama.FobjValorCampo(ldrwTotalDb(0), EnuTipoValor.enuDecimal)
        Return ldecTotalDb
    End Function
    Private Function FdecTotalCreditos(astrIdPredioAgr As String, adtmFechaTotal As Date) As Decimal 'Ok TipoNov
        Dim lstrFecha = "'" & ClsPanoramaDat.FstrFechaNormalizada(adtmFechaTotal) & "'"
        Dim lstrCamposSelect = {"SUM(" & ClsValor_NovDec.SstrNombreCampoBd & ")"}
        Dim lstbFiltro As New StringBuilder
        With lstbFiltro
            .Clear().Append(ClsOrionCop.StrFiltroUbicacion).Append(" AND ")
            .Append(ClsIdTercero_NovDbl.SstrNombreCampoBd).Append(" = ").Append(ObjIdClienteDbl.ObjValorPro)
            .Append(" AND ").Append(ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd).Append(" = '")
            .Append(astrIdPredioAgr).Append("' AND (")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = ")
            .Append(EnuTipoNov.EnuRDbIvaInt).Append(" OR ")
            .Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" BETWEEN ").Append(EnuTipoNov.EnuCrPagoCap)
            .Append(" AND ").Append(EnuTipoNov.EnuCrRetCre)
            .Append(" OR ").Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" = ").Append(EnuTipoNov.EnuCrIvaGas)
            .Append(" OR ").Append(ClsIdTipoNovedadByt.SstrNombreCampoBd).Append(" BETWEEN ")
            .Append(EnuTipoNov.EnuRDbCap).Append(" AND ").Append(EnuTipoNov.EnuRDbInt)
            .Append(") AND ").Append(ClsFechaNovedadDtm.SstrNombreCampoBd).Append(" <= ").Append(lstrFecha)
        End With
        Dim lstrFiltro = lstbFiltro.ToString
        Dim ldtbTotalCr = ClsPanorama.FdtbDataTable(ClsNovedad.SstrNombreTabla, lstrCamposSelect, {{}},
                    lstrFiltro)
        Dim ldrwTotalCr As DataRow = ldtbTotalCr.Select()(0)
        Dim ldecTotalCr As Decimal = ClsPanorama.FobjValorCampo(ldrwTotalCr(0), EnuTipoValor.enuDecimal)
        Return ldecTotalCr
    End Function
    Private Function FdecSaldoAnticipoFecha(astrIdPredioAgr As String, adtmFechaSaldo As Date) As Decimal
        Dim ldecSaldoAnticipo As Decimal
        Dim ldecTotalCr = FdecTotalCreditosAnt(astrIdPredioAgr, adtmFechaSaldo)
        Dim ldecTotalDb = FdecTotalDebitosAnt(astrIdPredioAgr, adtmFechaSaldo)
        ldecSaldoAnticipo = ldecTotalCr - ldecTotalDb
        Return ldecSaldoAnticipo
    End Function
    Private Function FdecTotalCreditosAnt(astrIdPredioAgr As String, adtmFechaTotal As Date) As Decimal 'Ok TipoNov
        Dim lstrFecha = "'" & ClsPanoramaDat.FstrFechaNormalizada(adtmFechaTotal) & "'"
        Dim lstrTablaPri = ClsNovedadAnticipo.SstrNombreTabla
        Dim lstrTablaSec = ClsAnticipo.SstrNombreTabla
        Dim lstrCamposSelectPri = {"SUM(" & ClsValor_NovAntDec.SstrNombreCampoBd & ")"}
        Dim lstrCamposSelectSec() As String = Array.Empty(Of String)
        Dim lstrCamposRelPri = {StrCampoCarpeta,
                                StrCampoCentroUtil,
                                ClsIdAnticipo_NovEnt.SstrNombreCampoBd}
        Dim lstrCamposRelSec = {StrCampoCarpeta,
                                StrCampoCentroUtil,
                                ClsIdAnticipoEnt.SstrNombreCampoBd}
        Dim lstbFiltro As New StringBuilder
        Dim lstrfiltro = ClsOrionCop.StrFiltroUbicacion_Pri
        With lstbFiltro
            .Clear().Append(lstrfiltro).Append(" AND ").Append("P.")
            .Append(ClsIdTercero_NovAntDbl.SstrNombreCampoBd)
            .Append(" = ").Append(ObjIdClienteDbl.ObjValorPro).Append(" AND ")
            .Append(ClsIdPredioAgrupador_AntStr.SstrNombreCampoBd).Append(" = '")
            .Append(astrIdPredioAgr).Append("'").Append(" AND (")
            .Append(ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd).Append(" = ")
            .Append(EnuTipoNov.EnuCrAntRec).Append(" OR ")
            .Append(ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd).Append(" BETWEEN ")
            .Append(EnuTipoNov.EnuRDbAntDev).Append(" AND ")
            .Append(EnuTipoNov.EnuRDbAntApl).Append(") AND ")
            .Append(ClsFechaNovedadAntDtm.SstrNombreCampoBd).Append(" <= ")
            .Append(lstrFecha)
        End With
        lstrfiltro = lstbFiltro.ToString
        Dim ldtbTotalCr = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCamposSelectPri,
                lstrTablaSec, lstrCamposSelectSec, lstrCamposRelPri, lstrCamposRelSec,
                {{}}, lstrfiltro, Array.Empty(Of String), True)
        Dim ldrwTotal As DataRow = ldtbTotalCr.Rows(0)
        Dim ldecTotalCr As Decimal = ClsPanorama.FobjValorCampo(ldrwTotal(0), EnuTipoValor.enuDecimal)
        Return ldecTotalCr
    End Function
    Private Function FdecTotalDebitosAnt(astrIdPredioAgr As String, adtmFechaTotal As Date) As Decimal 'Ok TipoNov
        Dim lstrFecha = "'" & ClsPanoramaDat.FstrFechaNormalizada(adtmFechaTotal) & "'"
        Dim lstrTablaPri = ClsNovedadAnticipo.SstrNombreTabla
        Dim lstrTablaSec = ClsAnticipo.SstrNombreTabla
        Dim lstrCamposSelectPri = {"SUM(" & ClsValor_NovAntDec.SstrNombreCampoBd & ")"}
        Dim lstrCamposSelectSec() As String = Array.Empty(Of String)
        Dim lstrCamposRelPri = {StrCampoCarpeta, StrCampoCentroUtil,
                                ClsIdAnticipo_NovEnt.SstrNombreCampoBd}
        Dim lstrCamposRelSec = {StrCampoCarpeta, StrCampoCentroUtil,
                                ClsIdAnticipoEnt.SstrNombreCampoBd}
        Dim lstbFiltro As New StringBuilder
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion_Pri
        With lstbFiltro
            .Clear().Append(lstrFiltro).Append(" AND ").Append("P.")
            .Append(ClsIdTercero_NovAntDbl.SstrNombreCampoBd).Append(" = ")
            .Append(ObjIdClienteDbl.ObjValorPro).Append(" AND ")
            .Append(ClsIdPredioAgrupador_AntStr.SstrNombreCampoBd)
            .Append(" = '").Append(astrIdPredioAgr).Append("'").Append(" AND (")
            .Append(ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd).Append(" BETWEEN ")
            .Append(EnuTipoNov.EnuDbAntDev).Append(" AND ").Append(EnuTipoNov.EnuDbAntApl)
            .Append(" OR ").Append(ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd).Append(" = ").Append(EnuTipoNov.EnuRCrAntRec)
            .Append(") AND ").Append(ClsFechaNovedadAntDtm.SstrNombreCampoBd).Append(" <= ").Append(lstrFecha)
        End With
        lstrFiltro = lstbFiltro.ToString
        Dim ldtbTotalDb = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCamposSelectPri, lstrTablaSec,
                lstrCamposSelectSec, lstrCamposRelPri, lstrCamposRelSec, {{}}, lstrFiltro,
                Array.Empty(Of String), True)
        Dim ldrwTotal As DataRow = ldtbTotalDb.Rows(0)
        Dim ldecTotalDb As Decimal = ClsPanorama.FobjValorCampo(ldrwTotal(0), EnuTipoValor.enuDecimal)
        Return ldecTotalDb
    End Function
    ''' <summary>
    ''' Devuelve la totalidad de intereses de mora por causar de los predios agrupadores pasados
    ''' en el argumento
    ''' </summary>
    ''' <param name="astrIdPrediosAgru"></param>
    ''' <param name="adtmFecha"></param>
    ''' <returns></returns>
    Friend Function FdecIntMoraPorCausar(astrIdPrediosAgru As String(), adtmFecha As Date) As Decimal
        Dim ldecIntMora = 0D
        Dim ldtbFacturasVivas = FdtbIdFras(astrIdPrediosAgru, {"A"}, True)
        Dim lstrPref As String, lentIdFac As Integer
        Dim lobjFac As New ClsFactura()
        For Each ldrwIdFac As DataRow In ldtbFacturasVivas.Rows
            lstrPref = ClsPanorama.FobjValorCampo(ldrwIdFac(
                    ClsPrefijo_ItemFactStr.SstrNombreCampoBd), EnuTipoValor.enuString)
            lentIdFac = ClsPanorama.FobjValorCampo(ldrwIdFac(
                    ClsIdFactura_ItemFactEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
            lobjFac.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentIdFac})
            ldecIntMora += lobjFac.FdecIntMoraPorCausar(adtmFecha)
        Next
        Return ldecIntMora
    End Function

    ''' <summary>
    ''' Devueleve un ArrayList con la identificación de las facturas con saldo del Cliente 
    ''' para el predio agrupador "astrIdPredioAgru" y los servicios incluidos en el
    ''' arreglo astrServicios
    ''' "astrServicios"
    ''' </summary>
    Friend Function FstrIdFacturasVivas(astrIdPredioAgru As String,
             astrServicios As String()) As ArrayList
        Dim lstrIdFactVivas As New ArrayList
        Dim ldtbIdFrasVivas = FdtbIdFras({astrIdPredioAgru}, astrServicios, True)
        Dim lstrPref As String, lentIdFac As Integer, lstrNroFac As String
        For Each ldrwIdFac As DataRow In ldtbIdFrasVivas.Rows
            lstrPref = ClsPanorama.FobjValorCampo(ldrwIdFac(0), EnuTipoValor.EnuString)
            lentIdFac = ClsPanorama.FobjValorCampo(ldrwIdFac(1), EnuTipoValor.EnuInteger)
            lstrNroFac = ClsPanorama.FstrNumeroDcto(lstrPref, lentIdFac)
            lstrIdFactVivas.Add(lstrNroFac)
        Next
        Return lstrIdFactVivas
    End Function

    ''' <summary>
    ''' Devuelve un ArrayList con la identificación de las facturas de los predios agrupadores
    ''' pasados en el argumento, que tienen deuda por intereses de mora
    ''' </summary>
    ''' <param name="astrIdPrediosAgru"></param>
    ''' <returns></returns>
    Friend Function FstrIdFacturasConMora(astrIdPrediosAgru As String()) As ArrayList 'Ok
        Dim lstrIdFactVivas As New ArrayList
        Dim lcolFrasVivas = FcolFacturas(astrIdPrediosAgru, {}, True)
        For Each lobjFactura As ClsFactura In lcolFrasVivas
            If lobjFactura.FdecDeudaIntTotal > 0 Then
                lstrIdFactVivas.Add(lobjFactura.StrNumeroFactura)
            End If
        Next
        Return lstrIdFactVivas
    End Function

    ''' <summary>
    ''' Devueleve un ArrayList con la identificación de las facturas del Cliente con saldo  
    ''' para los predios agrupadores en "astrIdPrediosAgru" y los servicios incluidos en astrServicios
    ''' </summary>
    ''' <param name="astrIdPrediosAgru">Array con los Predios Agrupadores a los cuales pertenecen 
    ''' las facturas.</param>
    ''' <param name="astrServicios">Identifica al los servicios incluidos en
    ''' las facturas</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function FstrIdFacturasVivas(astrIdPrediosAgru As String(),
             astrServicios As String()) As ArrayList
        Dim lstrIdFactVivas As New ArrayList
        Dim lcolFrasVivas = FcolFacturas(astrIdPrediosAgru, astrServicios, True)
        For Each lobjFactura As ClsFactura In lcolFrasVivas
            If lobjFactura.FdecDeudaCapitalSer(astrServicios) > 0 Then
                lstrIdFactVivas.Add(lobjFactura.StrNumeroFactura)
            End If
        Next
        Return lstrIdFactVivas
    End Function

    ''' <summary>
    ''' Devuelve un DataTable con la información de las facturas con saldo pertenecientes a los Predios Agrupadores 
    ''' indicados en el parametro "astrIdPrediosAgru" y al Agrupador de Servicios indicado en el 
    ''' parametro "ashrIdAgruServicios"
    ''' </summary>
    Friend Function FdtbDetalleDeuda(astrIdPrediosAgru As String(),
            astrServicios As String()) As DataTable
        GobjPanDat.SControleProcesoObj(True)
        Dim lcolFrasVivas = FcolFacturas(astrIdPrediosAgru, astrServicios, True)
        If IsNothing(MdtbFacturasDeuda) Then
            MdtbFacturasDeuda = ClsOrionCop.FdtbDetalleDeuda
        End If
        MdtbFacturasDeuda.Clear()
        Dim ldrwNuevoDetalle As DataRow, i = 0, ldecDeudaCap As Decimal, ldecTotalDeuda = 0D
        Dim ldecDeudaMora = FdecDeudaMora(astrIdPrediosAgru)
        If ldecDeudaMora > 0 Then
            i += 1
            ldecTotalDeuda += ldecDeudaMora
            ldrwNuevoDetalle = MdtbFacturasDeuda.NewRow
            ldrwNuevoDetalle("Ordinal") = i
            ldrwNuevoDetalle("NroFactura") = "Todas"
            ldrwNuevoDetalle("FechaFact") = DateSerial(1900, 1, 1)
            ldrwNuevoDetalle("FechaVence") = Date.Today
            ldrwNuevoDetalle("ValorFactura") = 0
            ldrwNuevoDetalle("DeudaCapital") = 0
            ldrwNuevoDetalle("DeudaIntMoras") = ldecDeudaMora
            ldrwNuevoDetalle("DeudaTotal") = ldecTotalDeuda
            MdtbFacturasDeuda.Rows.Add(ldrwNuevoDetalle)
        End If
        Dim lstrNroFac As String
        For Each lobjFact As ClsFactura In lcolFrasVivas
            i += 1
            lstrNroFac = lobjFact.StrNumeroFactura
            ldecDeudaCap = lobjFact.FdecDeudaCapitalSer(astrServicios)
            If ldecDeudaCap > 0 Then
                ldecTotalDeuda += ldecDeudaCap
                ldrwNuevoDetalle = MdtbFacturasDeuda.NewRow
                ldrwNuevoDetalle("Ordinal") = i
                ldrwNuevoDetalle("NroFactura") = lstrNroFac
                ldrwNuevoDetalle("FechaFact") = lobjFact.ObjFechaFacturaDtm.ObjValorPro
                ldrwNuevoDetalle("FechaVence") = lobjFact.ObjFechaVencimientoDtm.ObjValorPro
                ldrwNuevoDetalle("ValorFactura") = lobjFact.ObjValor_FactDec.ObjValorPro
                ldrwNuevoDetalle("DeudaCapital") = ldecDeudaCap
                ldrwNuevoDetalle("DeudaIntMoras") = 0
                ldrwNuevoDetalle("DeudaTotal") = ldecTotalDeuda
                MdtbFacturasDeuda.Rows.Add(ldrwNuevoDetalle)
            End If
        Next
        GobjPanDat.SControleProcesoObj(False)
        Return MdtbFacturasDeuda
    End Function
    ''' <summary>
    ''' Devuelve el estado sugerido del cliente incluyendo las facturas de todos los predios agrupadores
    '''  del cliente
    ''' </summary>
    Friend Function FenuEstadoSugeridoDeuda() As EnuEstadoDeudaDef
        GobjPanDat.SControleProcesoObj(True)
        Dim ldblIdCliente = 0.0
        If ObjIdClienteDbl.ObjValorPro > 0 Then
            ldblIdCliente = ObjIdClienteDbl.ObjValorPro
        End If
        Dim lenuEstadoSugeridoCliente = EnuEstadoDeudaDef.EnuNormal
        Dim lstrCamposSelect = {"MIN(" & ClsFechaVencimientoDtm.SstrNombreCampoBd & ")"}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsDebitos_FactDec.SstrNombreCampoBd &
                " <> " & ClsCreditos_FactDec.SstrNombreCampoBd & " AND " &
                ClsIdCliente_FactDbl.SstrNombreCampoBd & " = " & ldblIdCliente
        Dim ldtbMinFecVen = ClsPanorama.FdtbDataTable(ClsFactura.SstrNombreTabla, lstrCamposSelect,
                {{"", ""}}, lstrFiltro)
        Dim ldtmFechaVen = GCDTMFECHANULA
        If ldtbMinFecVen.Rows.Count > 0 Then
            If Not IsDBNull(ldtbMinFecVen.Select()(0)(0)) Then
                ldtmFechaVen = ClsPanorama.FobjValorCampo(ldtbMinFecVen.Select()(0)(0), EnuTipoValor.EnuDate)
            End If
        End If
        If ldtmFechaVen <> GCDTMFECHANULA Then
            Dim lentDiasVen = DateDiff(DateInterval.Day, ldtmFechaVen, Date.Today)
            With GobjParametros
                If lentDiasVen >= .ObjDiasParaPerdidaShr.ObjValorPro Then
                    lenuEstadoSugeridoCliente = EnuEstadoDeudaDef.EnuPerdida
                ElseIf lentDiasVen < .ObjDiasParaPerdidaShr.ObjValorPro AndAlso
                        lentDiasVen >= .ObjDiasParaJuridicoShr.ObjValorPro Then
                    lenuEstadoSugeridoCliente = EnuEstadoDeudaDef.EnuJuridico
                ElseIf lentDiasVen < .ObjDiasParaJuridicoShr.ObjValorPro AndAlso
                        lentDiasVen >= .ObjDiasParaPrejuridicoShr.ObjValorPro Then
                    lenuEstadoSugeridoCliente = EnuEstadoDeudaDef.EnuPrejuridico
                ElseIf lentDiasVen < .ObjDiasParaPrejuridicoShr.ObjValorPro AndAlso
                        lentDiasVen >= .ObjDiasParaPersuasivoShr.ObjValorPro Then
                    lenuEstadoSugeridoCliente = EnuEstadoDeudaDef.EnuPersuasivo
                Else
                    lenuEstadoSugeridoCliente = EnuEstadoDeudaDef.EnuNormal
                End If
            End With
        End If
        GobjPanDat.SControleProcesoObj(False)
        Return lenuEstadoSugeridoCliente
    End Function
    Friend Function FenuEstadoSugeridoDeuda(astrIdPredioAgrupador As String)
        GobjPanDat.SControleProcesoObj(True)
        Dim ldblIdCliente = 0.0
        If Not IsNothing(ObjIdClienteDbl.ObjValorPro) Then
            ldblIdCliente = CType(ObjIdClienteDbl.ObjValorPro, Double)
        End If
        Dim lenuEstadoSugeridoCliente = EnuEstadoDeudaDef.EnuNormal
        Dim lstrCamposSelect = {"MIN(" & ClsFechaVencimientoDtm.SstrNombreCampoBd & ")"}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsDebitos_FactDec.SstrNombreCampoBd &
                " <> " & ClsCreditos_FactDec.SstrNombreCampoBd & " AND " &
                ClsIdCliente_FactDbl.SstrNombreCampoBd & " = " & ldblIdCliente & " AND " &
                ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd & " = '" & astrIdPredioAgrupador & "'"
        Dim ldtbMinFecVen = ClsPanorama.FdtbDataTable(ClsFactura.SstrNombreTabla, lstrCamposSelect,
                {{"", ""}}, lstrFiltro)
        Dim ldtmFechaVen = GCDTMFECHANULA
        If ldtbMinFecVen.Rows.Count > 0 Then
            If Not IsDBNull(ldtbMinFecVen.Select()(0)(0)) Then
                ldtmFechaVen = ClsPanorama.FobjValorCampo(ldtbMinFecVen.Select()(0)(0), EnuTipoValor.EnuDate)
            End If
        End If
        If ldtmFechaVen <> GCDTMFECHANULA Then
            Dim lentDiasVen = DateDiff(DateInterval.Day, ldtmFechaVen, Date.Today)
            With GobjParametros
                If lentDiasVen >= .ObjDiasParaPerdidaShr.ObjValorPro Then
                    lenuEstadoSugeridoCliente = EnuEstadoDeudaDef.EnuPerdida
                ElseIf lentDiasVen < .ObjDiasParaPerdidaShr.ObjValorPro AndAlso
                        lentDiasVen >= .ObjDiasParaJuridicoShr.ObjValorPro Then
                    lenuEstadoSugeridoCliente = EnuEstadoDeudaDef.EnuJuridico
                ElseIf lentDiasVen < .ObjDiasParaJuridicoShr.ObjValorPro AndAlso
                        lentDiasVen >= .ObjDiasParaPrejuridicoShr.ObjValorPro Then
                    lenuEstadoSugeridoCliente = EnuEstadoDeudaDef.EnuPrejuridico
                ElseIf lentDiasVen < .ObjDiasParaPrejuridicoShr.ObjValorPro AndAlso
                        lentDiasVen >= .ObjDiasParaPersuasivoShr.ObjValorPro Then
                    lenuEstadoSugeridoCliente = EnuEstadoDeudaDef.EnuPersuasivo
                Else
                    lenuEstadoSugeridoCliente = EnuEstadoDeudaDef.EnuNormal
                End If
            End With
        End If
        GobjPanDat.SControleProcesoObj(False)
        Return lenuEstadoSugeridoCliente
    End Function
    Friend Sub SCambieEstadoDeuda(aenuEstadoDeuda As EnuEstadoDeudaDef, astrIdPredioAgrupador As String)
        If astrIdPredioAgrupador <> String.Empty Then
            SCambieEstadoDeudaPre(aenuEstadoDeuda, astrIdPredioAgrupador)
        Else
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando Then
                ObjIdEstadoDeudaByt.ObjValorPro = aenuEstadoDeuda
                SActualice(True)
            End If
        End If
    End Sub
    Private Shared Sub SCambieEstadoDeudaPre(aenuEstadoDeuda As EnuEstadoDeudaDef, astrIdPredioAgrupador As String)
        Dim lobjPredioAgr As New ClsPredio(EnuModoInstanciaObjDef.EnuUnico)
        lobjPredioAgr.SAbra({GshrIdCarpeta, GshrIdCentroUtil, astrIdPredioAgrupador})
        For Each lobjPredio As ClsPredio In lobjPredioAgr.ColPrediosAgrupados
            lobjPredio.SModifique()
            If lobjPredio.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando Then
                lobjPredio.ObjIdEstadoDeuda_PredioByt.ObjValorPro = aenuEstadoDeuda
                lobjPredio.SActualice(True)
            End If
        Next
    End Sub
    Friend Function FenuEstadoDeuda(astrIdPredioAgr As String)
        Dim lenuEstadoDeuda As EnuEstadoDeudaDef = EnuEstadoDeudaDef.None
        Dim lstrIdPrediosAgr As String() = astrIdPredioAgr.Split(",")
        For Each lstrIdPredioAgr As String In lstrIdPrediosAgr
            If String.IsNullOrEmpty(lstrIdPredioAgr) Then
                If ObjIdEstadoDeudaByt.ObjValorPro > lenuEstadoDeuda Then
                    lenuEstadoDeuda = ObjIdEstadoDeudaByt.ObjValorPro
                End If
            Else
                Dim lobjPredioAgr As New ClsPredio(EnuModoInstanciaObjDef.EnuUnico)
                lobjPredioAgr.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrIdPredioAgr})
                If lobjPredioAgr.ObjIdEstadoDeuda_PredioByt.ObjValorPro > lenuEstadoDeuda Then
                    lenuEstadoDeuda = lobjPredioAgr.ObjIdEstadoDeuda_PredioByt.ObjValorPro
                End If
            End If
        Next
        Return lenuEstadoDeuda
    End Function

    Friend Function FobjEstadoCtaHoy(astrIdPredioAgr As String, astrSer As String,
            ByRef adtbItemsFacEstado As DataTable)
        MobjEstadoCtaHoy = Nothing
        Dim ldblIdClie = ObjIdClienteDbl.ObjValorPro
        MobjEstadoCtaHoy = If(String.IsNullOrEmpty(astrIdPredioAgr),
                ClsOrionCop.FobjEstadoCtaHoy(ldblIdClie, astrSer, adtbItemsFacEstado),
                ClsOrionCop.FobjEstadoCtaHoy(ldblIdClie, astrIdPredioAgr, astrSer,
                        adtbItemsFacEstado))
        If adtbItemsFacEstado.Rows.Count > 0 Then
            SComplementeDtbEstado(adtbItemsFacEstado, astrSer)
        End If
        Return MobjEstadoCtaHoy
    End Function

    Private Sub SComplementeDtbEstado(adtbItemsFacEstado As DataTable, astrSer As String)
        Dim ldclVlrItem As New DataColumn("Valor", Type.GetType("System.Decimal"))
        Dim ldclDeuCap As New DataColumn("DeudaCapital", Type.GetType("System.Decimal"))
        Dim ldclDeuMor As New DataColumn("DeudaMora", Type.GetType("System.Decimal"))
        Dim ldclSaldo As New DataColumn("Saldo", Type.GetType("System.Decimal"))
        adtbItemsFacEstado.Columns.Add(ldclVlrItem)
        adtbItemsFacEstado.Columns.Add(ldclDeuCap)
        adtbItemsFacEstado.Columns.Add(ldclDeuMor)
        adtbItemsFacEstado.Columns.Add(ldclSaldo)
        Dim lstrPref As String, lentIdFac As Integer, lstrIdFact As String
        Dim lstrIdFacs As New List(Of String)
        Dim ldecVlrItem As Decimal, ldecDeuCap As Decimal, ldecDeuMora As Decimal
        Dim ldecSaldo As Decimal, ldrwItemsRemov As New List(Of DataRow)
        Dim lobjFac As New ClsFactura()
        For Each ldrwTbl As DataRow In adtbItemsFacEstado.Rows
            lstrPref = ClsPanorama.FobjValorCampo(ldrwTbl(
                    ClsPrefijo_ItemFactStr.SstrNombreCampoBd), EnuTipoValor.EnuString)
            lentIdFac = ClsPanorama.FobjValorCampo(ldrwTbl(
                    ClsIdFactura_ItemFactEnt.SstrNombreCampoBd), EnuTipoValor.EnuInteger)
            lstrIdFact = lstrPref & "-" & lentIdFac.ToString()
            If Not lstrIdFacs.Contains(lstrIdFact) Then
                lstrIdFacs.Add(lstrIdFact)
                lobjFac.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentIdFac})
                ldecVlrItem = lobjFac.FdecCapitalServiciosFact({astrSer})
                ldecDeuCap = lobjFac.FdecDeudaCapitalSer({astrSer})
                ldecDeuMora = lobjFac.FdecDeudaIntMoraSer({astrSer})
                ldecSaldo = ldecDeuCap + ldecDeuMora
                ldrwTbl("Valor") = ldecVlrItem
                ldrwTbl("DeudaCapital") = ldecDeuCap
                ldrwTbl("DeudaMora") = ldecDeuMora
                ldrwTbl("Saldo") = ldecSaldo
            Else
                ldrwItemsRemov.Add(ldrwTbl)
            End If
        Next
        For Each ldrwRem As DataRow In ldrwItemsRemov
            adtbItemsFacEstado.Rows.Remove(ldrwRem)
        Next
    End Sub
#End Region

#Region "Facturas del Cliente"
    Friend ReadOnly Property ObjNuevaFactura(aenuModoFacturacion As EnuModoFacturacionDef) As ClsFactura
        Get
            Dim ldrwNuevaFactura As DataRow = FdrwNuevoDataRowFactura()
            Dim lobjNuevaFactura As New ClsFactura(Me, ldrwNuevaFactura)
            With lobjNuevaFactura
                .SCreeObj(Nothing)
                .ObjIdCliente_FactDbl.ObjValorPro = Me.ObjIdClienteDbl.ObjValorPro
                .ObjIdModoFacturacionByt.ObjValorPro = aenuModoFacturacion
            End With
            Return lobjNuevaFactura
        End Get
    End Property

    Friend Function FdtbServiciosConDeuda(astrPredAgru As String()) As DataTable
        Dim lstrFiltro = " AND (", lblnUltimo As Boolean, i = 0
        For Each lstrPreAgr As String In astrPredAgru
            lstrFiltro &= ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd &
                    " = '" & lstrPreAgr
            lblnUltimo = i = astrPredAgru.GetUpperBound(0)
            If lblnUltimo Then
                lstrFiltro &= "')"
            Else
                lstrFiltro &= "' OR "
            End If
            i += 1
        Next
        Dim lsbExp = New StringBuilder
        With lsbExp
            .Append("SELECT DISTINCT ").Append(ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd)
            .Append(", ").Append(ClsIdServicio_ItemFactShr.SstrNombreCampoBd).Append(" FROM ")
            .Append(ClsItemFactura.SstrNombreTabla).Append(" AS I INNER JOIN ")
            .Append(ClsFactura.SstrNombreTabla).Append(" AS F ON I.")
            .Append(ClsPrefijo_ItemFactStr.SstrNombreCampoBd).Append(" = F.")
            .Append(ClsPrefijo_FactStr.SstrNombreCampoBd).Append(" AND I.")
            .Append(ClsIdFactura_ItemFactEnt.SstrNombreCampoBd).Append(" = F.")
            .Append(ClsIdFacturaEnt.SstrNombreCampoBd).Append(" WHERE I.").Append(StrCampoCarpeta)
            .Append(" = ").Append(GshrIdCarpeta).Append(" AND I.")
            .Append(StrCampoCentroUtil).Append(" = ").Append(GshrIdCentroUtil)
            .Append(" AND ").Append(ClsIdCliente_FactDbl.SstrNombreCampoBd).Append(" = ")
            .Append(ObjIdClienteDbl.ToString).Append(lstrFiltro).Append(" AND I.")
            .Append(ClsDebitos_ItemFactDec.SstrNombreCampoBd).Append(" - I.")
            .Append(ClsCreditos_ItemFactDec.SstrNombreCampoBd).Append(" > 0").Append(" ORDER BY ")
            .Append(ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd).Append(" DESC ,")
            .Append(ClsIdServicio_ItemFactShr.SstrNombreCampoBd)
        End With
        Dim ldtbAgrSer = ClsPanorama.FdtbDataTable(lsbExp.ToString)
        Return ldtbAgrSer
    End Function
    Private Function FdrwNuevoDataRowFactura() As DataRow
        Dim ldrwNuevo As DataRow
        If IsNothing(MdtbFacturas) Then
            Dim lstrIndice = {{StrCampoCarpeta, "ASC"},
                              {StrCampoCentroUtil, "ASC"},
                              {ClsIdCliente_FactDbl.SstrNombreCampoBd, "ASC"},
                              {ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd, "ASC"},
                              {ClsFechaFacturaDtm.SstrNombreCampoBd, "ASC"},
                              {ClsPrefijo_FactStr.SstrNombreCampoBd, "ASC"},
                              {ClsIdFacturaEnt.SstrNombreCampoBd, "ASC"}}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsIdCliente_FactDbl.SstrNombreCampoBd & " = 0"
            Dim ldtbFacturas As DataTable = ClsPanorama.FdtbDataTable(ClsFactura.SstrNombreTabla, {"*"},
                               lstrIndice, lstrFiltro)
            ldrwNuevo = ldtbFacturas.NewRow
        Else
            ldrwNuevo = MdtbFacturas.NewRow
        End If
        Return ldrwNuevo
    End Function
    Friend Function FdtmFechaUltimaNovedad(astrIdPredioAgrupador As String)
        Dim ldtmFechaUltNov As Date
        Dim lstrTablaPri = ClsNovedad.SstrNombreTabla
        Dim lstrTablaSec = ClsFactura.SstrNombreTabla
        Dim lstrCamSelPri = {"MAX(" & ClsFechaNovedadDtm.SstrNombreCampoBd & ")"}
        Dim lstrCamSelSec = Array.Empty(Of String)
        Dim lstrCamRelPri = {ClsPrefijoFact_NovStr.SstrNombreCampoBd, ClsIdFactura_NovEnt.SstrNombreCampoBd,
                ClsIdTercero_NovDbl.SstrNombreCampoBd, ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd}
        Dim lstrCamRelSec = {ClsPrefijo_FactStr.SstrNombreCampoBd, ClsIdFacturaEnt.SstrNombreCampoBd,
                ClsIdCliente_FactDbl.SstrNombreCampoBd, ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd}
        Dim lstrOrden(,) As String = {{}}
        Dim lstrfiltro = ClsOrionCop.StrFiltroUbicacion_Pri & " AND P." &
                ClsIdTercero_NovDbl.SstrNombreCampoBd & " = " & ObjIdClienteDbl.ToString &
                " AND P." & ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd & " = '" & astrIdPredioAgrupador &
                "' AND " & ClsDebitos_FactDec.SstrNombreCampoBd & " > " &
                ClsCreditos_FactDec.SstrNombreCampoBd & " AND S." & ClsAnuladoBln.SstrNombreCampoBd &
                " = FALSE"
        Dim ldtbMaxFechaNov As DataTable = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCamSelPri,
                lstrTablaSec, lstrCamSelSec, lstrCamRelPri, lstrCamRelSec, lstrOrden, lstrfiltro,
                Array.Empty(Of String), True)
        Dim ldrwMaxFechaNov As DataRow = ldtbMaxFechaNov.Select()(0)
        If IsDBNull(ldrwMaxFechaNov(0)) Then
            ldtmFechaUltNov = GCDTMFECHANULA
        Else
            ldtmFechaUltNov = ClsPanorama.FobjValorCampo(ldrwMaxFechaNov(0), EnuTipoValor.EnuDate)
        End If
        Return ldtmFechaUltNov
    End Function
    Friend Function FarlPrediosAgrDelCliente() As ArrayList
        Dim larlPredagrCli As New ArrayList, lstrPredAgr As String
        Dim ldtbPredsAgrClie = FdtbPrediosAgrupDelCliEnFac()
        Dim ldtbPredsAgrPropCli = FdtbPredAgruPropios()
        For Each ldrwPredAgrCli As DataRow In ldtbPredsAgrClie.Rows
            lstrPredAgr = ClsPanorama.FobjValorCampo(ldrwPredAgrCli(
                    ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd), EnuTipoValor.EnuString)
            If Not larlPredagrCli.Contains(lstrPredAgr) Then
                larlPredagrCli.Add(lstrPredAgr)
            End If
        Next
        For Each ldrwPredAgrCli As DataRow In ldtbPredsAgrPropCli.Rows
            lstrPredAgr = ClsPanorama.FobjValorCampo(ldrwPredAgrCli(
                    ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd), EnuTipoValor.EnuString)
            If Not larlPredagrCli.Contains(lstrPredAgr) Then
                larlPredagrCli.Add(lstrPredAgr)
            End If
        Next
        larlPredagrCli.Sort()
        Return larlPredagrCli
    End Function
    Private Function FdtbPrediosAgrupDelCliEnFac() As DataTable
        Dim lstbSql As New StringBuilder
        With lstbSql
            .Append("SELECT DISTINCT ").Append(ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd)
            .Append(" FROM ").Append(ClsFactura.SstrNombreTabla).Append(" WHERE ")
            .Append(StrCampoCarpeta).Append(" = ").Append(GshrIdCarpeta)
            .Append(" AND ").Append(StrCampoCentroUtil).Append(" = ")
            .Append(GshrIdCentroUtil).Append(" AND ").Append(ClsIdCliente_FactDbl.SstrNombreCampoBd)
            .Append(" = ").Append(ObjIdClienteDbl.ToString).Append(" ORDER BY ")
            .Append(ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd).Append(" ASC")
        End With
        Dim lstrSql = lstbSql.ToString
        Dim ldtbPrediosAgrup = ClsPanorama.FdtbDataTable(lstrSql)
        Return ldtbPrediosAgrup
    End Function
    Private Function FdtbPredAgruPropios() As DataTable
        Dim lstrTablaPri = ClsPredio.SstrNombreTabla
        Dim lstrTablaSec = ClsPropietario.SstrNombreTabla
        Dim lstrCampSelPri = {"DISTINCT " & ClsIdPredioAgrupadorStr.SstrNombreCampoBd}
        Dim lstrCampSelSec As String() = {}
        Dim lstrCampRelPri = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsIdPredioStr.SstrNombreCampoBd}
        Dim lstrCampRelSec = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsIdPredio_PropStr.SstrNombreCampoBd}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion_Pri & " AND " & "P." &
                ClsIdPredioStr.SstrNombreCampoBd & " = " &
                ClsIdPredioAgrupadorStr.SstrNombreCampoBd & " AND S." &
                ClsIdCliente_PropDbl.SstrNombreCampoBd & " = " & ObjIdClienteDbl.ToString()
        Dim lstrOrden As String(,) = {{ClsIdPredioAgrupadorStr.SstrNombreCampoBd, "ASC"}}
        Dim ldtbPreAgrProp = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCampSelPri,
                lstrTablaSec, lstrCampSelSec, lstrCampRelPri, lstrCampRelSec, lstrOrden, True,
                lstrFiltro, {})
        Return ldtbPreAgrProp
    End Function
    Friend Function FarlPrediosAgrupEnFras() As ArrayList
        Dim larlIdPrediosAgrup As New ArrayList
        Dim lstrCamposSelect() = {"DISTINCT " & ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd}
        Dim lstrIndice(,) = {{ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdCliente_FactDbl.SstrNombreCampoBd & " = " & ObjIdClienteDbl.ObjValorPro
        Dim ldtbPrediosAgrup = ClsPanorama.FdtbDataTable(ClsFactura.SstrNombreTabla, lstrCamposSelect,
                                                         lstrIndice, lstrFiltro)
        Dim ldrwPrediosAgru() As DataRow = ldtbPrediosAgrup.Select()
        For i = 0 To ldrwPrediosAgru.Count - 1
            larlIdPrediosAgrup.Add(ClsPanorama.FobjValorCampo(ldrwPrediosAgru(i)(0), EnuTipoValor.EnuString))
        Next
        Dim larlPreGruEnRC = FarlPrediosAgrupEnRC()
        For Each lstrIdPreAgr In larlPreGruEnRC
            If Not larlIdPrediosAgrup.Contains(lstrIdPreAgr) Then
                larlIdPrediosAgrup.Add(lstrIdPreAgr)
            End If
        Next
        Return larlIdPrediosAgrup
    End Function
    Friend Function FarlPrediosAgrupEnRC() As ArrayList
        Dim larlIdPrediosAgrup As New ArrayList
        Dim lstrCamposSelect() = {"DISTINCT " & ClsIdPredioAgrupador_RecStr.SstrNombreCampoBd}
        Dim lstrIndice(,) = {{ClsIdPredioAgrupador_RecStr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdCliente_FactDbl.SstrNombreCampoBd & " = " & ObjIdClienteDbl.ToString()
        Dim ldtbPrediosAgrup = ClsPanorama.FdtbDataTable(ClsReciboCaja.SstrNombreTabla, lstrCamposSelect,
                                                         lstrIndice, lstrFiltro)
        Dim ldrwPrediosAgru() As DataRow = ldtbPrediosAgrup.Select()
        Dim lstrIdPrediosAgr As String()
        For i = 0 To ldrwPrediosAgru.Count - 1
            lstrIdPrediosAgr = ClsPanorama.FobjValorCampo(ldrwPrediosAgru(i)(0),
                    EnuTipoValor.EnuString).ToString.Split(",")
            For Each lstrIdpreAgr As String In lstrIdPrediosAgr
                If Not larlIdPrediosAgrup.Contains(lstrIdpreAgr) Then
                    larlIdPrediosAgrup.Add(lstrIdpreAgr)
                End If
            Next
        Next
        Return larlIdPrediosAgrup
    End Function
    Friend Function FarlPrediosAgrupEnNCR() As ArrayList
        Dim larlIdPrediosAgrup As New ArrayList
        Dim lstrCamposSelect() = {"DISTINCT " & ClsIdPredioAgrupador_NotaCrStr.SstrNombreCampoBd}
        Dim lstrIndice(,) = {{ClsIdPredioAgrupador_NotaCrStr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdCliente_FactDbl.SstrNombreCampoBd & " = " & ObjIdClienteDbl.ObjValorPro
        Dim ldtbPrediosAgrup = ClsPanorama.FdtbDataTable(ClsNotaCr.SstrNombreTabla, lstrCamposSelect,
                                                         lstrIndice, lstrFiltro)
        Dim ldrwPrediosAgru() As DataRow = ldtbPrediosAgrup.Select()
        For i = 0 To ldrwPrediosAgru.Count - 1
            larlIdPrediosAgrup.Add(ClsPanorama.FobjValorCampo(ldrwPrediosAgru(i)(0), EnuTipoValor.EnuString))
        Next
        Return larlIdPrediosAgrup
    End Function
    Friend Function FarlPrediosAgrupEnNDB() As ArrayList
        Dim larlIdPrediosAgrup As New ArrayList
        Dim lstrCamposSelect() = {"DISTINCT " & ClsIdPredioAgrupador_NotaDbStr.SstrNombreCampoBd}
        Dim lstrIndice(,) = {{ClsIdPredioAgrupador_NotaDbStr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdCliente_FactDbl.SstrNombreCampoBd & " = " & ObjIdClienteDbl.ObjValorPro
        Dim ldtbPrediosAgrup = ClsPanorama.FdtbDataTable(ClsNotaDb.SstrNombreTabla, lstrCamposSelect,
                                                         lstrIndice, lstrFiltro)
        Dim ldrwPrediosAgru() As DataRow = ldtbPrediosAgrup.Select()
        For i = 0 To ldrwPrediosAgru.Count - 1
            larlIdPrediosAgrup.Add(ClsPanorama.FobjValorCampo(ldrwPrediosAgru(i)(0), EnuTipoValor.EnuString))
        Next
        Return larlIdPrediosAgrup
    End Function
    Friend Function FarlPrediosAgrupEnNAA() As ArrayList
        Dim larlIdPrediosAgrup As New ArrayList
        Dim lstrCamposSelect() = {"DISTINCT " & ClsIdPredioAgrupador_NotaConStr.SstrNombreCampoBd}
        Dim lstrIndice(,) = {{ClsIdPredioAgrupador_NotaConStr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdCliente_FactDbl.SstrNombreCampoBd & " = " & ObjIdClienteDbl.ObjValorPro
        Dim ldtbPrediosAgrup = ClsPanorama.FdtbDataTable(ClsNotaCon.SstrNombreTabla, lstrCamposSelect,
                                                         lstrIndice, lstrFiltro)
        Dim ldrwPrediosAgru() As DataRow = ldtbPrediosAgrup.Select()
        For i = 0 To ldrwPrediosAgru.Count - 1
            larlIdPrediosAgrup.Add(ClsPanorama.FobjValorCampo(ldrwPrediosAgru(i)(0), EnuTipoValor.EnuString))
        Next
        Return larlIdPrediosAgrup
    End Function
    Friend Function FblnEsFactElec(astrPref As String, aentIdFac As Integer) As Boolean
        Dim lblnEsFacEle = GobjParametros.BlnEFacAutorizado
        If lblnEsFacEle Then
            Dim lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, astrPref, aentIdFac}
            Dim lobjFac As New ClsFactura()
            lobjFac.SAbra(lobjValorLlave)
            lblnEsFacEle = lobjFac.BlnEsFacEle
        End If
        Return lblnEsFacEle
    End Function
    ''' <summary>
    ''' Indica si una factura del cliente está en estado registrada o enviada
    ''' </summary>
    ''' <returns></returns>
    Friend Function FblnFactEstadoEFacOk(astrPref As String, aentIdFac As Integer) As Boolean
        Dim lblnFacEstadoOk = False
        If GobjParametros.BlnEFacAutorizado Then
            Dim lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, astrPref, aentIdFac}
            Dim lobjFac As New ClsFactura()
            lobjFac.SAbra(lobjValorLlave)
            lblnFacEstadoOk = lobjFac.ObjIdEstadoEDocEnt.ObjValorPro >= EnuEstadoEDoc.EnuRegi
        End If
        Return lblnFacEstadoOk
    End Function
    ''' <summary>
    ''' Indica si todas las facturas electrónivas vivas del cliente estan en estado Ok 
    ''' </summary>
    ''' <returns></returns>
    Friend Function FblnEstadoEFacOk(astrIdPredioAgr As String) As Boolean
        Dim lblnOk = True
        If GobjParametros.BlnEFacAutorizado Then
            Dim lstrTabla = ClsFactura.SstrNombreTabla
            Dim lstrCamSel As String() = {ClsIdFacturaEnt.SstrNombreCampoBd}
            Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsIdCliente_FactDbl.SstrNombreCampoBd & " = " & ObjIdClienteDbl.ObjValorPro &
                    " AND " & ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd & " = '" &
                    astrIdPredioAgr & "'" & " AND " & ClsIdEstadoEDocEnt.SstrNombreCampoBd & " < " &
                    EnuEstadoEDoc.EnuRegi
            Dim lstrOrden As String(,) = {{"", ""}}
            Dim ldtbRes = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamSel, lstrOrden, lstrFiltro)
            lblnOk = Not ldtbRes.Rows.Count > 0
        End If
        Return lblnOk
    End Function
#End Region

#Region "Items Programa Facturacion"
    Friend ReadOnly Property ColItemsProgramaFact() As Collection
        Get
            If IsNothing(McolItemsProgramaFact) Then
                McolItemsProgramaFact = New Collection
                SCargueDtbItemsProgFact()
                Dim ldrwItemsprogramaFact() As DataRow = MdtbItemsProgFact.Select()
                If ldrwItemsprogramaFact.Length > 0 Then
                    Dim lshrIdAno As Short
                    Dim lstrKey As String
                    For Each ldrwItemProgFac As DataRow In ldrwItemsprogramaFact
                        lshrIdAno = ClsPanorama.FobjValorCampo(ldrwItemProgFac(
                                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd), EnuTipoValor.EnuShort)
                        If lshrIdAno = 0 Then
                            Dim lobjItemprogramaFact As New ClsItemProgramaFact(Me,
                                    EnuTipoDeudorDef.EnuCliente, ldrwItemProgFac)
                            lobjItemprogramaFact.SLeaValores(True)
                            lstrKey = lobjItemprogramaFact.ObjIdAno_ItemProgramaFactShr.ToString &
                                    "," & lobjItemprogramaFact.ObjIdServicio_ItemProgramaFactShr.ToString
                            McolItemsProgramaFact.Add(lobjItemprogramaFact, lstrKey)
                        Else
                            Throw New PanLException("Valor del año no valido en 'clsCliente.colItemsProgramaFact'")
                        End If
                    Next
                End If
            End If
            Return McolItemsProgramaFact
        End Get
    End Property
    Friend ReadOnly Property ObjNewItemProgramaFact As ClsItemProgramaFact
        Get
            Dim lobjItemProgramaFact As ClsItemProgramaFact = Nothing
            SCargueDtbItemsProgFact()
            If Not IsNothing(MdtbItemsProgFact) Then
                Dim ldrwNewItemProgramafact As DataRow = MdtbItemsProgFact.NewRow
                lobjItemProgramaFact = New ClsItemProgramaFact(Me, EnuTipoDeudorDef.EnuCliente,
                        ldrwNewItemProgramafact)
                With lobjItemProgramaFact
                    .SCreeObj(Nothing)
                    .ObjIdCarpeta_ItemProgramaFactShr.ObjValorPro = GshrIdCarpeta
                    .ObjIdCentroUtil_ItemProgramaFactShr.ObjValorPro = GshrIdCentroUtil
                    .ObjIdAno_ItemProgramaFactShr.ObjValorPro = 0
                    .ObjIdCliente_ItemProgramaFactDbl.ObjValorPro = ObjIdClienteDbl.ObjValorPro
                    .ObjIdPredio_ItemProgramaFactStr.ObjValorPro = String.Empty
                End With
            End If
            Return lobjItemProgramaFact
        End Get
    End Property
    Friend ReadOnly Property ObjItemProgramaFact(ashrIdServicio As Short) As ClsItemProgramaFact
        Get
            Dim lobjItemProgramaFact As ClsItemProgramaFact = Nothing
            SCargueDtbItemsProgFact()
            Dim ldrwItemsprogFact() As DataRow = MdtbItemsProgFact.Select
            For Each ldrwItemProFac As DataRow In ldrwItemsprogFact
                Dim lshrIdAno As Short = ClsPanorama.FobjValorCampo(
                        ldrwItemProFac(ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd), EnuTipoValor.EnuShort)
                If lshrIdAno = 0 Then
                    Dim lshrIdServicio As Short = ClsPanorama.FobjValorCampo(
                            ldrwItemProFac(ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd), EnuTipoValor.EnuShort)
                    If lshrIdServicio = ashrIdServicio Then
                        lobjItemProgramaFact = New ClsItemProgramaFact(Me, EnuTipoDeudorDef.EnuCliente,
                                ldrwItemProFac)
                        lobjItemProgramaFact.SLeaValores(True)
                        Exit For
                    End If
                Else
                    Throw New PanLException("Valor del año no valido en 'clsCliente.objItemProgramaFact'")
                End If
            Next
            Return lobjItemProgramaFact
        End Get
    End Property
    Friend ReadOnly Property DtbItemsProgFac() As DataTable
        Get
            MdtbItemsProgFact = Nothing
            SCargueDtbItemsProgFact()
            Return MdtbItemsProgFact
        End Get
    End Property
    Friend Sub SAgregueItemProgramaFact(aobjItemProgramaFact As ClsItemProgramaFact)
        If Not IsNothing(aobjItemProgramaFact) Then
            aobjItemProgramaFact.SActualice(True)
            Dim lstrKey = aobjItemProgramaFact.ObjIdAno_ItemProgramaFactShr.ToString &
                    "," & aobjItemProgramaFact.ObjIdServicio_ItemProgramaFactShr.ToString
            McolItemsProgramaFact.Add(aobjItemProgramaFact, lstrKey)
        End If
    End Sub
    Friend Function FblnPendienteporFacturar(astrKeyItemProgramaFact As String) As Boolean
        Dim lblnPorFacturar = False
        If ColItemsProgramaFact.Contains(astrKeyItemProgramaFact) Then
            Dim lobjItemProgramaFact As ClsItemProgramaFact = ColItemsProgramaFact(astrKeyItemProgramaFact)
            If Not IsNothing(lobjItemProgramaFact) Then
                lblnPorFacturar = lobjItemProgramaFact.FblnPendienteDeFacturar
            End If
        End If
        Return lblnPorFacturar
    End Function
    Private Sub SCargueDtbItemsProgFact()
        If IsNothing(MdtbItemsProgFact) Then
            Dim lstrIdCliente = ObjIdClienteDbl.ToString
            If String.IsNullOrEmpty(lstrIdCliente) Then lstrIdCliente = "0"
            GobjPanDat.SControleProcesoObj(True)
            Dim lstrCamposSelect() = {"*", "'*' AS NombreServicio", "'*' AS NombreOrigen"}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsIdCliente_ItemProgramaFactDbl.SstrNombreCampoBd & " = " & lstrIdCliente
            Dim lstrIndice = {{ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd, "ASC"},
                                 {ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd, "ASC"}}
            MdtbItemsProgFact = ClsPanorama.FdtbDataTable(ClsItemProgramaFact.SstrNombreTabla, lstrCamposSelect,
                    lstrIndice, lstrFiltro)
            SAsigneCamposComplemento()
            GobjPanDat.SControleProcesoObj(False)
        End If
    End Sub
    Private Sub SAsigneCamposComplemento()
        Dim ldrwItemsProgramaFact() As DataRow = MdtbItemsProgFact.Select
        If ldrwItemsProgramaFact.Count > 0 Then
            For Each ldrwItemProFac As DataRow In ldrwItemsProgramaFact
                Dim lshrIdAno As Short
                Dim lshrIdServicio As Short
                Dim lstrNombreServicio = String.Empty
                lshrIdServicio = ClsPanorama.FobjValorCampo(ldrwItemProFac(
                        ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd),
                            EnuTipoValor.EnuShort)
                lshrIdAno = ClsPanorama.FobjValorCampo(ldrwItemProFac(
                        ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd),
                            EnuTipoValor.EnuShort)
                Dim lstrKey = lshrIdAno.ToString & "," & lshrIdServicio.ToString
                Dim lobjServicio As ClsServicio = Nothing
                If lshrIdAno <> 0 Then
                    If GobjParametros.ColAnos.Contains(lshrIdAno.ToString) Then
                        Dim lobjAno As ClsAno = GobjParametros.ColAnos(lshrIdAno.ToString)
                        If lobjAno.ColServiciosAno.Contains(lstrKey) Then
                            lobjServicio = lobjAno.ColServiciosAno(lstrKey)
                        End If
                    End If
                Else
                    If GobjParametros.ColServiciosPer.Contains(lstrKey) Then
                        lobjServicio = GobjParametros.ColServiciosPer(lstrKey)
                    End If
                End If
                If Not IsNothing(lobjServicio) Then
                    lstrNombreServicio = lobjServicio.ObjNombreServicioStr.ObjValorPro
                End If
                ldrwItemProFac("NombreServicio") = lstrNombreServicio
                SAsigneOrigenEnDtb(ldrwItemProFac)
            Next
        End If
    End Sub
    Private Shared Sub SAsigneOrigenEnDtb(adrwItemOrigenProFac As DataRow)
        Dim lstrNombreCampoOrigen = ClsOrigen_ItemProgramaFacByt.SstrNombreCampoBd
        Dim lbytIdOrigen As Byte = ClsPanorama.FobjValorCampo(adrwItemOrigenProFac(lstrNombreCampoOrigen),
                    EnuTipoValor.EnuByte)
        Dim lstrNombreOrigen = ClsOrionCop.FstrNombreDatoConstanteOri(
                EnuGrupoConstantesOriDef.EnuOrigenItemProgramaFact, lbytIdOrigen)
        adrwItemOrigenProFac("NombreOrigen") = lstrNombreOrigen
    End Sub
#End Region

#Region "Notas Contables"
    Friend Function FobjNuevaNotaCon(adtmFecha As Date, astrIdpredioAgrupador As String) As ClsNotaCon
        SCargueDtbNotasCon(True)
        Dim ldrwNuevaNotaCon As DataRow = MdtbNotasCon.NewRow
        Dim lobjNuevaNotaCon As New ClsNotaCon(Me, ldrwNuevaNotaCon)
        Dim lblnCambioPermisos = False
        With lobjNuevaNotaCon
            If Not CType(.EnuPermisosObj And EnuPermisosDef.EnuCrear, Boolean) Then
                .EnuPermisosObj += EnuPermisosDef.EnuCrear
                lblnCambioPermisos = True
            End If
            .SCreeObj(Nothing)
            .ObjFecha_NotaConDtm.ObjValorPro = adtmFecha
            .ObjIdCliente_NotaConDbl.ObjValorPro = ObjIdClienteDbl.ObjValorPro
            .ObjIdPredioAgrupador_NotaConStr.ObjValorPro = astrIdpredioAgrupador
            If lblnCambioPermisos Then
                .EnuPermisosObj -= EnuPermisosDef.EnuCrear
            End If
        End With
        Return lobjNuevaNotaCon
    End Function
    Private Sub SCargueDtbNotasCon(ablnSoloEstructura As Boolean)
        Static lblnSoloEstructura As Boolean = True
        If IsNothing(MdtbNotasCon) OrElse (ablnSoloEstructura <> lblnSoloEstructura) Then
            lblnSoloEstructura = ablnSoloEstructura
            Dim lstrIndice = {{StrCampoCarpeta, "ASC"},
                              {StrCampoCentroUtil, "ASC"},
                              {ClsIdPredioAgrupador_NotaConStr.SstrNombreCampoBd, "ASC"},
                              {ClsPrefijo_NotaConStr.SstrNombreCampoBd, "ASC"},
                              {ClsIdNotaConEnt.SstrNombreCampoBd, "ASC"}}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsIdCliente_NotaConDbl.SstrNombreCampoBd & " = "
            If ablnSoloEstructura Then
                lstrFiltro &= "0"
            Else
                lstrFiltro &= ObjIdClienteDbl.ObjValorPro
            End If
            Dim lstrCamposSelect() = {"*"}
            MdtbNotasCon = ClsPanorama.FdtbDataTable(ClsNotaCon.SstrNombreTabla, lstrCamposSelect, lstrIndice,
                    lstrFiltro)
        End If
    End Sub
    Friend ReadOnly Property ObjNuevaNotaDevAnt() As ClsNotaDevAnt
        Get
            SCargueDtbNotasDevAnt()
            Dim ldrwNuevaNotaDevAnt As DataRow = MdtbNotasDevAnt.NewRow
            Dim lobjNuevaNotaDevAnt As New ClsNotaDevAnt(Me, ldrwNuevaNotaDevAnt)
            With lobjNuevaNotaDevAnt
                .SCreeObj(Nothing)
                .ObjFecha_NotaDevAntDtm.ObjValorPro = Date.Today
                .ObjFechaCreacionDtm.ObjValorPro = Date.Now
                .ObjIdAnticipo_NotaDevAntEnt.ObjValorPro = 0
                .ObjIdCarpeta_NotaDevAntShr.ObjValorPro = GshrIdCarpeta
                .ObjIdCentroUtil_NotaDevAntShr.ObjValorPro = GshrIdCentroUtil
                .ObjIdCliente_NotaDevAntDbl.ObjValorPro = ObjIdClienteDbl.ObjValorPro
                .ObjIdNotaDevAntEnt.ObjValorPro = 0
                .ObjIdPredioAgrupador_NotaDevAntStr.ObjValorPro = String.Empty
                .ObjIdUsuario_NotaDevAntStr.ObjValorPro = GstrIdUsuario
                .ObjOrigenInstanciaStr.ObjValorPro = GstrOrigenActual
                .ObjValor_NotaDevAntDec.ObjValorPro = 0
            End With
            Return lobjNuevaNotaDevAnt
        End Get
    End Property
    Private Sub SCargueDtbNotasDevAnt()
        If IsNothing(MdtbNotasDevAnt) Then
            GobjPanDat.SControleProcesoObj(True)
            Dim lstrCamposSelect() = {"*"}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsIdCliente_NotaDevAntDbl.SstrNombreCampoBd & " = " & ObjIdClienteDbl.ToString
            Dim lstrIndice = {{StrCampoCarpeta, "ASC"},
                              {StrCampoCentroUtil, "ASC"},
                              {ClsIdCliente_NotaDevAntDbl.SstrNombreCampoBd, "ASC"},
                              {ClsIdPredioAgrupador_NotaDevAntStr.SstrNombreCampoBd, "ASC"},
                              {ClsIdNotaDevAntEnt.SstrNombreCampoBd, "ASC"}}
            MdtbNotasDevAnt = ClsPanorama.FdtbDataTable(ClsNotaDevAnt.SstrNombreTabla, lstrCamposSelect,
                    lstrIndice, lstrFiltro)
            GobjPanDat.SControleProcesoObj(False)
        End If
    End Sub
#End Region

#Region "Notas Credito"
    Friend Function FdtbNotasCr(astrIdPredioAgr As String)
        SCargueDtbNotasCr(astrIdPredioAgr)
        Return MdtbNotasCr
    End Function
    Private Sub SCargueDtbNotasCr(astrIdPredioAgru As String)
        Static ldblIdCliente As Double = 0 : Static lstrIdPredioAgr As String = String.Empty
        If ldblIdCliente <> ObjIdClienteDbl.ObjValorPro OrElse lstrIdPredioAgr <> astrIdPredioAgru Then
            ldblIdCliente = ObjIdClienteDbl.ObjValorPro : lstrIdPredioAgr = astrIdPredioAgru
            GobjPanDat.SControleProcesoObj(True)
            Dim lstrAnu = "IIF(" & ClsAnuladoBln.SstrNombreCampoBd & " = 0, 'No', 'Si') AS Anu"
            Dim lstrCamposSelect() = {"*", lstrAnu}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsIdCliente_NotaCrDbl.SstrNombreCampoBd & " = " & ObjIdClienteDbl.ToString & " AND " &
                    ClsIdPredioAgrupador_NotaCrStr.SstrNombreCampoBd & " = '" & astrIdPredioAgru & "'"
            Dim lstrIndice = {{StrCampoCarpeta, "ASC"},
                              {StrCampoCentroUtil, "ASC"},
                              {ClsIdCliente_NotaCrDbl.SstrNombreCampoBd, "ASC"},
                              {ClsIdPredioAgrupador_NotaCrStr.SstrNombreCampoBd, "ASC"},
                              {ClsPrefijo_NotaCrStr.SstrNombreCampoBd, "ASC"},
                              {ClsIdNotaCrEnt.SstrNombreCampoBd, "ASC"}}
            MdtbNotasCr = ClsPanorama.FdtbDataTable(ClsNotaCr.SstrNombreTabla, lstrCamposSelect,
                    lstrIndice, lstrFiltro)
            GobjPanDat.SControleProcesoObj(False)
        End If
    End Sub
    Friend Function FdtbNotasCr(astrIdPredioAgr As String, adtmFechaDesde As Date,
             adtmFechaHasta As Date) As DataTable
        SCargueDtbNotasCr(astrIdPredioAgr, adtmFechaDesde, adtmFechaHasta)
        Return MdtbNotasCr
    End Function
    Private Sub SCargueDtbNotasCr(astrIdPredioAgru As String, ByRef adtmFechaDesde As Date,
                                    adtmFechaHasta As Date)
        Dim lstrExpSqlNotaCrNov = FstrExpSqlNotaCrNov(astrIdPredioAgru, adtmFechaDesde,
                adtmFechaHasta)
        Dim lstrIndice = ClsFecha_NotaCrDtm.SstrNombreCampoBd & ", " &
                  ClsPrefijo_NotaCrStr.SstrNombreCampoBd & ", " &
                  ClsIdNotaCrEnt.SstrNombreCampoBd
        Dim lstrExpSql = lstrExpSqlNotaCrNov & " ORDER BY " & lstrIndice
        GobjPanDat.SControleProcesoObj(True)
        MdtbNotasCr = ClsPanorama.FdtbDataTable(lstrExpSql)
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Private Function FstrExpSqlNotaCrNov(astrIdPredioAgru As String, ByRef adtmFechaDesde As Date,
             adtmFechaHasta As Date) As String
        Dim lstrFechaDesde = ClsPanoramaDat.FstrFechaNormalizada(adtmFechaDesde)
        Dim lstrFechaHasta = ClsPanoramaDat.FstrFechaNormalizada(adtmFechaHasta)
        Dim lstrTablaPri = ClsNovedad.SstrNombreTabla
        Dim lstrTablaSec = ClsNotaCr.SstrNombreTabla
        Dim lstrCamposPri() = {"DISTINCT " & ClsFechaNovedadDtm.SstrNombreCampoBd}
        Dim lstrCamposSec() = {"*"}
        Dim lstrCamposRelPri = {StrCampoCarpeta,
                                StrCampoCentroUtil,
                                ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd,
                                ClsIdDocOrigenEnt.SstrNombreCampoBd}
        Dim lstrCamposRelSec = {StrCampoCarpeta,
                                StrCampoCentroUtil,
                                ClsPrefijo_NotaCrStr.SstrNombreCampoBd,
                                ClsIdNotaCrEnt.SstrNombreCampoBd}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion_Pri & " AND " &
                ClsIdTipoDocOrigenByt.SstrNombreCampoBd & " = " & EnuTipoDocOri.EnuNotaCr &
                " AND S." & ClsIdCliente_NotaCrDbl.SstrNombreCampoBd & " = " &
                ObjIdClienteDbl.ToString & " AND P." &
                ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd & " = '" & astrIdPredioAgru & "'"
        lstrFiltro &= " AND " & ClsFecha_NotaCrDtm.SstrNombreCampoBd & " >= '" & lstrFechaDesde &
                "' AND " & ClsFecha_NotaCrDtm.SstrNombreCampoBd & " <='" & lstrFechaHasta & "'"
        Dim lstrSql As String = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTablaPri, lstrCamposPri,
            lstrTablaSec, lstrCamposSec, lstrCamposRelPri, lstrCamposRelSec,
            {{"", ""}}, lstrFiltro, Array.Empty(Of String))
        Return lstrSql
    End Function
#End Region

#Region " Notas Débito"
    ''' <summary>
    ''' Esta colección contiene las notas debito del cliente. Si el argumento "astrIdPredioAgru" es Null
    ''' contiene todas las notas debito, de lo contrario contiene las notas debito del predio agrupador
    ''' identificado por el argumento.
    ''' </summary>
    ''' <param name="astrIdPredioAgru">Id del predio agrupador al cualpertenecen las notas Db.</param>
    ''' <value></value>
    Friend ReadOnly Property ColNotasDb(astrIdPredioAgru As String) As Collection
        Get
            If IsNothing(McolNotasDb) OrElse McolNotasDb.Count = 0 OrElse
                    astrIdPredioAgru <> MstrPredioAgruNotasDb Then
                MstrPredioAgruNotasDb = astrIdPredioAgru
                McolNotasDb = New Collection
                SCargueDtbNotasDb(False)
                If Not IsNothing(MdtbNotasDb) AndAlso MdtbNotasDb.Rows.Count > 0 Then
                    Dim lstrFiltro = String.Empty
                    If Not IsNothing(astrIdPredioAgru) Then
                        lstrFiltro &= ClsIdPredioAgrupador_NotaDbStr.SstrNombreCampoBd & " = '" & astrIdPredioAgru
                    End If
                    Dim ldrwNotasDb() As DataRow = MdtbNotasDb.Select(lstrFiltro)
                    For Each ldrwNotaDb As DataRow In ldrwNotasDb
                        Dim lobjNotaDb As New ClsNotaDb(Me, ldrwNotaDb)
                        lobjNotaDb.SLeaValores(True)
                        McolNotasDb.Add(lobjNotaDb, lobjNotaDb.ObjIdNotaDbEnt.ToString)
                    Next
                End If
            End If
            Return McolNotasDb
        End Get
    End Property
    Friend Function FdtbNotasDb(astrIdPredioAgr As String, adtmFechaDesde As Date,
             adtmFechaHasta As Date)
        MdtbNotasDbExtendida = FdtbNotasDbPredioAgr(astrIdPredioAgr, adtmFechaDesde, adtmFechaHasta)
        Return MdtbNotasDbExtendida
    End Function
    Friend Function FobjNuevaNotaDb(adtmFecha As Date, astrIdpredioAgrupador As String,
            aenuOrigen As EnuOrigenNotaDb) As ClsNotaDb
        SCargueDtbNotasDb(True)
        Dim ldrwNuevaNotaDb As DataRow = MdtbNotasDb.NewRow
        Dim lobjNuevaNotaDb As New ClsNotaDb(Me, ldrwNuevaNotaDb)
        With lobjNuevaNotaDb
            .EnuPermisosObj += EnuPermisosDef.EnuCrear
            .SCreeObj(Nothing)
            .ObjAnuladoBln.ObjValorPro = False
            .ObjFechaAnulacion_NotaDbDtm.ObjValorPro = GCDTMFECHANULA
            .ObjFechaCreacionDtm.ObjValorPro = Date.Now
            .ObjIdCarpeta_NotaDbShr.ObjValorPro = GshrIdCarpeta
            .ObjIdCentroUtil_NotaDbShr.ObjValorPro = GshrIdCentroUtil
            .ObjIdCliente_NotaDbDbl.ObjValorPro = ObjIdClienteDbl.ObjValorPro
            .ObjIdUsuario_NotaDbStr.ObjValorPro = GstrIdUsuario
            .ObjIdUsuarioAnuloStr.ObjValorPro = String.Empty
            .ObjOrigenInstanciaStr.ObjValorPro = GstrOrigenActual
            .ObjOrigenInstanciaAnuloStr.ObjValorPro = String.Empty
            .ObjPrefijo_NotaDbStr.ObjValorPro = GobjParametros.FstrPrefijoDoc(
                    EnuTipoDocOri.EnuNotaDb)
            .ObjFecha_NotaDbDtm.ObjValorPro = adtmFecha
            .ObjIdPredioAgrupador_NotaDbStr.ObjValorPro = astrIdpredioAgrupador
            .ObjOrigenByt.ObjValorPro = aenuOrigen
        End With
        Return lobjNuevaNotaDb
    End Function
    Friend Function FobjNuevaNotaDb(astrIdpredioAgrupador As String,
            aenuOrigen As EnuOrigenNotaDb) As ClsNotaDb
        SCargueDtbNotasDb(True)
        Dim ldrwNuevaNotaDb As DataRow = MdtbNotasDb.NewRow
        Dim lobjNuevaNotaDb As New ClsNotaDb(Me, ldrwNuevaNotaDb)
        With lobjNuevaNotaDb
            .EnuPermisosObj += EnuPermisosDef.EnuCrear
            .SCreeObj(Nothing)
            .ObjAnuladoBln.ObjValorPro = False
            .ObjFechaAnulacion_NotaDbDtm.ObjValorPro = GCDTMFECHANULA
            .ObjFechaCreacionDtm.ObjValorPro = Date.Now
            .ObjIdCarpeta_NotaDbShr.ObjValorPro = GshrIdCarpeta
            .ObjIdCentroUtil_NotaDbShr.ObjValorPro = GshrIdCentroUtil
            .ObjIdCliente_NotaDbDbl.ObjValorPro = ObjIdClienteDbl.ObjValorPro
            .ObjIdUsuario_NotaDbStr.ObjValorPro = GstrIdUsuario
            .ObjIdUsuarioAnuloStr.ObjValorPro = String.Empty
            .ObjOrigenInstanciaStr.ObjValorPro = GstrOrigenActual
            .ObjOrigenInstanciaAnuloStr.ObjValorPro = String.Empty
            .ObjPrefijo_NotaDbStr.ObjValorPro = GobjParametros.FstrPrefijoDoc(
                    EnuTipoDocOri.EnuNotaDb)
            .ObjIdPredioAgrupador_NotaDbStr.ObjValorPro = astrIdpredioAgrupador
            .ObjOrigenByt.ObjValorPro = aenuOrigen
        End With
        Return lobjNuevaNotaDb
    End Function
    Private Function FdtbNotasDbPredioAgr(astrIdPredioAgr As String, adtmFechaDesde As Date,
             adtmFechaHasta As Date) As DataTable
        GobjPanDat.SControleProcesoObj(True)
        Dim lstrFechaDesde = ClsPanoramaDat.FstrFechaNormalizada(adtmFechaDesde)
        Dim lstrFechaHasta = ClsPanoramaDat.FstrFechaNormalizada(adtmFechaHasta)
        Dim lstrAnu = "IIF(" & ClsAnuladoBln.SstrNombreCampoBd & " = 0, 'No', 'Si') AS Anu"
        Dim lstrCamposSelect = {"*", lstrAnu}
        Dim lstrIndice = {{StrCampoCarpeta, "ASC"},
                          {StrCampoCentroUtil, "ASC"},
                          {ClsIdPredioAgrupador_NotaDbStr.SstrNombreCampoBd, "ASC"},
                          {ClsPrefijo_NotaDbStr.SstrNombreCampoBd, "ASC"},
                          {ClsIdNotaDbEnt.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdCliente_NotaDbDbl.SstrNombreCampoBd & " = " & ObjIdClienteDbl.ObjValorPro &
                " AND " & ClsIdPredioAgrupador_NotaDbStr.SstrNombreCampoBd & " = '" & astrIdPredioAgr
        lstrFiltro &= "' AND " & ClsFecha_NotaDbDtm.SstrNombreCampoBd & " >= '" & lstrFechaDesde &
                "' AND " & ClsFecha_NotaDbDtm.SstrNombreCampoBd & " <='" & lstrFechaHasta & "'"
        Dim ldtbNotasDb = ClsPanorama.FdtbDataTable(ClsNotaDb.SstrNombreTabla, lstrCamposSelect,
                lstrIndice, lstrFiltro)
        GobjPanDat.SControleProcesoObj(False)
        Return ldtbNotasDb
    End Function
    Private Sub SCargueDtbNotasDb(ablnSoloEstructura As Boolean)
        If IsNothing(MdtbNotasDb) OrElse (ablnSoloEstructura <> MblnCargueDtbSoloEstrucura) Then
            MblnCargueDtbSoloEstrucura = ablnSoloEstructura
            Dim lstrIndice = {{StrCampoCarpeta, "ASC"},
                              {StrCampoCentroUtil, "ASC"},
                              {ClsIdPredioAgrupador_NotaDbStr.SstrNombreCampoBd, "ASC"},
                              {ClsPrefijo_NotaDbStr.SstrNombreCampoBd, "ASC"},
                              {ClsIdNotaDbEnt.SstrNombreCampoBd, "ASC"}}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsIdCliente_NotaDbDbl.SstrNombreCampoBd & " = "
            If ablnSoloEstructura Then
                lstrFiltro &= "0"
            Else
                lstrFiltro &= ObjIdClienteDbl.ObjValorPro
            End If
            Dim lstrCamposSelect() = {"*"}
            MdtbNotasDb = ClsPanorama.FdtbDataTable(ClsNotaDb.SstrNombreTabla, lstrCamposSelect, lstrIndice,
                    lstrFiltro)
        End If
    End Sub
#End Region

#Region "Notas Ajuste"
    Friend Function FobjNuevaNotaAjuste() As ClsNotaAjusteCuotaAdmin
        SCargueDtbNotasAjuste(True)
        Dim ldrwNuevaNotaAjuste As DataRow = MdtbNotasAjuste.NewRow
        Dim lobjNuevaNotaAjuste As New ClsNotaAjusteCuotaAdmin(Me, ldrwNuevaNotaAjuste)
        Dim lblnModificoPermisos = False
        With lobjNuevaNotaAjuste
            If Not CType(.EnuPermisosObj And EnuPermisosDef.EnuCrear, Boolean) Then
                .EnuPermisosObj += EnuPermisosDef.EnuCrear
                lblnModificoPermisos = True
            End If
            .SCreeObj(Nothing)
            .ObjIdCliente_NotaAjusteDbl.ObjValorPro = ObjIdClienteDbl.ObjValorPro
            If lblnModificoPermisos Then
                .EnuPermisosObj -= EnuPermisosDef.EnuCrear
            End If
        End With
        Return lobjNuevaNotaAjuste
    End Function
    Private Sub SCargueDtbNotasAjuste(ablnSoloEstructura As Boolean)
        If IsNothing(MdtbNotasAjuste) Then
            Dim lstrIndice = {{StrCampoCarpeta, "ASC"},
                              {StrCampoCentroUtil, "ASC"},
                              {ClsIdPredioAgrupador_NotaAjusteStr.SstrNombreCampoBd, "ASC"},
                              {ClsIdNotaAjusteEnt.SstrNombreCampoBd, "ASC"}}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsIdCliente_NotaAjusteDbl.SstrNombreCampoBd & " = "
            If ablnSoloEstructura Then
                lstrFiltro &= "0"
            Else
                lstrFiltro &= ObjIdClienteDbl.ObjValorPro
            End If
            Dim lstrCamposSelect() = {"*"}
            MdtbNotasAjuste = ClsPanorama.FdtbDataTable(ClsNotaAjusteCuotaAdmin.SstrNombreTabla, lstrCamposSelect,
                    lstrIndice, lstrFiltro)
        End If
    End Sub
#End Region

#Region "Novedades"
    ''' <summary>
    ''' Devuelve un datatable con el mvimiento en novedades complementado con el saldo inicial,
    ''' el detalle y el saldo para un predio agrupador dado entre dos fechas
    ''' </summary>
    Friend Function FdtbMovimientoDeuda(astrIdPredioAgr As String, adtmFechaDesde As Date,
             adtmFechaHasta As Date)
        GobjPanDat.SControleProcesoObj(True)
        SCargueDtbNovedadesDeuda(astrIdPredioAgr, adtmFechaDesde, adtmFechaHasta)
        SComplementeDtbMovimientoDeuda(astrIdPredioAgr, adtmFechaDesde)
        GobjPanDat.SControleProcesoObj(False)
        Return MdtbNovedadesDeuda
    End Function
    Private Sub SCargueDtbNovedadesDeuda(astrIdPredioAgru As String, ByRef adtmFechaDesde As Date,
                                    adtmFechaHasta As Date)
        Dim lstrFechaDesde = "'" & ClsPanoramaDat.FstrFechaNormalizada(adtmFechaDesde) & "'"
        Dim lstrFechaHasta = "'" & ClsPanoramaDat.FstrFechaNormalizada(adtmFechaHasta) & "'"
        Dim lstrExpSqlNov = FstrExprNovedadesDeuda(astrIdPredioAgru, lstrFechaDesde, lstrFechaHasta)
        ' 
        MdtbNovedadesDeuda = ClsPanorama.FdtbDataTable(lstrExpSqlNov)
    End Sub
    Private Function FstrExprNovedadesDeuda(astrIdPredioAgru As String, ByRef astrFechaDesde As String,
                                    astrFechaHasta As String) As String
        Dim lstrCamposSelect() = {ClsFechaCreacionDtm.SstrNombreCampoBd,
                                  ClsFechaNovedadDtm.SstrNombreCampoBd,
                                  ClsIdTercero_NovDbl.SstrNombreCampoBd,
                                  ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd,
                                  ClsIdTipoDocOrigenByt.SstrNombreCampoBd,
                                  ClsIdTipoNovedadByt.SstrNombreCampoBd,
                                  ("CONCAT(" & ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd & ",'-' , " &
                                   ClsIdDocOrigenEnt.SstrNombreCampoBd & ")") & " AS NroDocOri",
                                  "SUM(" & ClsValor_NovDec.SstrNombreCampoBd & ") AS Valor",
                                  "'' AS Detalle", "0 AS Saldo"}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdTercero_NovDbl.SstrNombreCampoBd & " = " & ObjIdClienteDbl.ToString & " AND " &
                ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd & " = '" & astrIdPredioAgru & "'"
        lstrFiltro &= " AND " & ClsFechaNovedadDtm.SstrNombreCampoBd & " >= " & astrFechaDesde &
                " AND " & ClsFechaNovedadDtm.SstrNombreCampoBd & " <=" & astrFechaHasta
        Dim lstrIndice = {{ClsFechaCreacionDtm.SstrNombreCampoBd, " ASC"}}
        Dim lstrCamposGrupo = {ClsFechaNovedadDtm.SstrNombreCampoBd,
                               ClsIdTipoDocOrigenByt.SstrNombreCampoBd,
                               ClsIdTipoNovedadByt.SstrNombreCampoBd, " NroDocOri",
                               " Detalle", " Saldo"}
        Dim lstrSqlNov = ClsPanoramaDat.FstrConstruyaExpSqlSelect(ClsNovedad.SstrNombreTabla, lstrCamposSelect,
                lstrIndice, lstrFiltro, lstrCamposGrupo)
        Return lstrSqlNov
    End Function
    Private Sub SComplementeDtbMovimientoDeuda(astrIdPredioAgr As String, adtmFechaDesde As Date) 'Ok TipoNov
        Dim ldecSaldo = FdecSaldoDeudaFecha(astrIdPredioAgr, adtmFechaDesde.AddDays(-1))
        Dim ldrwSaldoInicial = MdtbNovedadesDeuda.NewRow
        ldrwSaldoInicial(ClsFechaNovedadDtm.SstrNombreCampoBd) = adtmFechaDesde.AddDays(-1)
        ldrwSaldoInicial(ClsIdTercero_NovDbl.SstrNombreCampoBd) = ObjIdClienteDbl.ObjValorPro
        ldrwSaldoInicial(ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd) = astrIdPredioAgr
        ldrwSaldoInicial("Detalle") = "Saldo inicial a la fecha"
        ldrwSaldoInicial("Saldo") = ldecSaldo
        MdtbNovedadesDeuda.Rows.InsertAt(ldrwSaldoInicial, 0)
        Dim ldrwRegistro As DataRow, lstrDetalle As String, ldecValor As Decimal
        Dim lenuTipoNovedad As EnuTipoNov
        For i = 1 To MdtbNovedadesDeuda.Rows.Count - 1
            ldrwRegistro = MdtbNovedadesDeuda.Rows(i)
            Dim lstrIdDoc As String = ClsPanorama.FobjValorCampo(ldrwRegistro("NroDocOri"),
                    EnuTipoValor.EnuString)
            If lstrIdDoc.StartsWith(" -") Then
                lstrIdDoc = lstrIdDoc.Substring(1)
            End If
            ldrwRegistro("NroDocOri") = lstrIdDoc
            ldecValor = ClsPanorama.FobjValorCampo(ldrwRegistro(ClsValor_NovDec.SstrNombreCampoBd),
                    EnuTipoValor.EnuDecimal)
            lenuTipoNovedad = ClsPanorama.FobjValorCampo(ldrwRegistro(ClsIdTipoNovedadByt.SstrNombreCampoBd),
                    EnuTipoValor.EnuByte)
            If (lenuTipoNovedad >= EnuTipoNov.EnuCrPagoCap AndAlso
                    lenuTipoNovedad <= EnuTipoNov.EnuCrRetCre) OrElse
                    (lenuTipoNovedad >= EnuTipoNov.EnuRDbCap AndAlso
                    lenuTipoNovedad <= EnuTipoNov.EnuRDbInt) OrElse
                    lenuTipoNovedad = EnuTipoNov.EnuCrIvaGas OrElse
                    lenuTipoNovedad = EnuTipoNov.EnuRDbIvaInt Then
                ldecSaldo -= ldecValor
            ElseIf (lenuTipoNovedad >= EnuTipoNov.EnuDbCap AndAlso
                    lenuTipoNovedad <= EnuTipoNov.EnuDbInt) OrElse
                    (lenuTipoNovedad >= EnuTipoNov.EnuRCrPagoCap AndAlso
                     lenuTipoNovedad <= EnuTipoNov.EnuRCrRetCre) OrElse
                     lenuTipoNovedad = EnuTipoNov.EnuDbIvaInt Then
                ldecSaldo += ldecValor
            Else
                Throw New ErrorInesperadoPanLException("Tipo de Novedad movimiento no esperada")
            End If
            lstrDetalle = FstrDetalle(ldrwRegistro, False)
            ldrwRegistro("Detalle") = lstrDetalle
            ldrwRegistro("Saldo") = ldecSaldo
        Next
    End Sub
    ''' <summary>
    ''' Devuelve un datatable con el movimiento de anticipos complementado con el saldo inicial,
    ''' el detalle y el saldo para un predio agrupador dado entre dos fechas
    ''' </summary>
    Friend Function FdtbMovimientoAnticipos(astrIdPredioAgr As String, adtmFechaDesde As Date,
             adtmFechaHasta As Date)
        GobjPanDat.SControleProcesoObj(True)
        SCargueDtbNovedadesAnt(astrIdPredioAgr, adtmFechaDesde, adtmFechaHasta)
        SComplementeDtbMovimientoAnt(astrIdPredioAgr, adtmFechaDesde)
        GobjPanDat.SControleProcesoObj(False)
        Return MdtbNovedadesAnticipo
    End Function
    Private Sub SCargueDtbNovedadesAnt(astrIdPredioAgru As String, ByRef adtmFechaDesde As Date,
                                    adtmFechaHasta As Date)
        Dim lstrFechaDesde = "'" & ClsPanoramaDat.FstrFechaNormalizada(adtmFechaDesde) & "'"
        Dim lstrFechaHasta = "'" & ClsPanoramaDat.FstrFechaNormalizada(adtmFechaHasta) & "'"
        Dim lstrExpSqlNov = FstrExprNovedadesAnt(astrIdPredioAgru, lstrFechaDesde, lstrFechaHasta)
        MdtbNovedadesAnticipo = ClsPanorama.FdtbDataTable(lstrExpSqlNov)
    End Sub
    Private Function FstrExprNovedadesAnt(astrIdPredioAgru As String, ByRef astrFechaDesde As String,
                                    astrFechaHasta As String) As String
        Dim lstrTablaPri = ClsNovedadAnticipo.SstrNombreTabla
        Dim lstrTablaSec = ClsAnticipo.SstrNombreTabla
        Dim lstrCamposSelectPri() = {ClsFechaNovedadAntDtm.SstrNombreCampoBd,
                                     ClsIdTercero_NovAntDbl.SstrNombreCampoBd,
                                     ClsIdTipoDocOrigen_AntByt.SstrNombreCampoBd,
                                     ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd,
                                     ClsPrefijoDocOrigen_NovAntStr.SstrNombreCampoBd,
                                     ClsIdDocOrigen_NovAntEnt.SstrNombreCampoBd,
                                     ClsValor_NovAntDec.SstrNombreCampoBd, "'' AS Detalle", "0 AS Saldo"}
        Dim lstrCamposSelectSec() As String = {ClsIdPredioAgrupador_AntStr.SstrNombreCampoBd}
        Dim lstrCamposRelPri() As String = {StrCampoCarpeta,
                                  StrCampoCentroUtil,
                                  ClsIdAnticipo_NovEnt.SstrNombreCampoBd}
        Dim lstrCamposRelSec() As String = {StrCampoCarpeta,
                                 StrCampoCentroUtil,
                                 ClsIdAnticipoEnt.SstrNombreCampoBd}
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion_Pri & " AND P." &
                ClsIdTercero_NovAntDbl.SstrNombreCampoBd & " = " &
                ObjIdClienteDbl.ToString & " AND " &
                ClsIdPredioAgrupador_AntStr.SstrNombreCampoBd & " = '" &
                astrIdPredioAgru & "'"
        lstrFiltro &= " AND " & ClsFechaNovedadAntDtm.SstrNombreCampoBd & " >= " &
                astrFechaDesde & " AND " & ClsFechaNovedadAntDtm.SstrNombreCampoBd &
                " <= " & astrFechaHasta
        Dim lstrOrden(,) As String = {{ClsFechaNovedadAntDtm.SstrNombreCampoBd, "ASC"},
                                     {ClsIdDocOrigen_NovAntEnt.SstrNombreCampoBd, "ASC"}}
        Dim lstrSqlNov = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTablaPri,
                lstrCamposSelectPri, lstrTablaSec, lstrCamposSelectSec, lstrCamposRelPri,
                lstrCamposRelSec, lstrOrden, lstrFiltro, Array.Empty(Of String))
        Return lstrSqlNov
    End Function
    Private Sub SComplementeDtbMovimientoAnt(astrIdPredioAgr As String, adtmFechaDesde As Date) 'Ok TipoNov
        Dim ldecSaldo = FdecSaldoAnticipoFecha(astrIdPredioAgr, adtmFechaDesde.AddDays(-1))
        Dim ldrwSaldoInicial = MdtbNovedadesAnticipo.NewRow
        ldrwSaldoInicial(ClsFechaNovedadDtm.SstrNombreCampoBd) = adtmFechaDesde.AddDays(-1)
        ldrwSaldoInicial(ClsIdTercero_NovDbl.SstrNombreCampoBd) = ObjIdClienteDbl.ObjValorPro
        ldrwSaldoInicial(ClsIdPredioAgrupador_AntStr.SstrNombreCampoBd) = astrIdPredioAgr
        ldrwSaldoInicial("Detalle") = "Saldo inicial a la fecha"
        ldrwSaldoInicial("Saldo") = ldecSaldo
        MdtbNovedadesAnticipo.Rows.InsertAt(ldrwSaldoInicial, 0)
        Dim ldrwRegistro As DataRow, lstrDetalle As String, ldecValor As Decimal
        Dim lenuTipoNovedadAnt As EnuTipoNov
        For i = 1 To MdtbNovedadesAnticipo.Rows.Count - 1
            ldrwRegistro = MdtbNovedadesAnticipo.Rows(i)
            ldecValor = ClsPanorama.FobjValorCampo(ldrwRegistro(ClsValor_NovAntDec.SstrNombreCampoBd),
                    EnuTipoValor.EnuDecimal)
            lenuTipoNovedadAnt = ClsPanorama.FobjValorCampo(ldrwRegistro(ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd),
                    EnuTipoValor.EnuByte)
            Select Case lenuTipoNovedadAnt
                Case EnuTipoNov.EnuCrAntRec, EnuTipoNov.EnuRDbAntDev, EnuTipoNov.EnuRDbAntApl
                    ldecSaldo += ldecValor
                Case EnuTipoNov.EnuDbAntApl, EnuTipoNov.EnuDbAntDev, EnuTipoNov.EnuRCrAntRec
                    ldecSaldo -= ldecValor
                Case Else
                    Throw New ErrorInesperadoPanLException("Tipo de Novedad en Anticipos no esperada")
            End Select
            lstrDetalle = FstrDetalle(ldrwRegistro, True)
            ldrwRegistro("Detalle") = lstrDetalle
            ldrwRegistro("Saldo") = ldecSaldo
        Next
    End Sub
    Private Shared Function FstrDetalle(adrwNovedad As DataRow, ablnAnticipo As Boolean) As String 'Ok TipoNov
        Dim lstrDocOri As String
        Dim lenuTipodocOrigen As EnuTipoDocOri
        Dim lenuTipoNovedad As EnuTipoNov
        If ablnAnticipo Then
            Dim lstrPref As String = ClsPanorama.FobjValorCampo(adrwNovedad(
                    ClsPrefijoDocOrigen_NovAntStr.SstrNombreCampoBd), EnuTipoValor.EnuString)
            Dim lentIdDoc As Integer = ClsPanorama.FobjValorCampo(adrwNovedad(
                    ClsIdDocOrigen_NovAntEnt.SstrNombreCampoBd), EnuTipoValor.EnuInteger)
            lstrDocOri = ClsPanorama.FstrNumeroDcto(lstrPref, lentIdDoc)
            lenuTipodocOrigen = adrwNovedad(ClsIdTipoDocOrigen_AntByt.SstrNombreCampoBd)
            lenuTipoNovedad = adrwNovedad(ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd)
        Else
            lstrDocOri = ClsPanorama.FobjValorCampo(adrwNovedad("NroDocOri"), EnuTipoValor.EnuString)
            lenuTipodocOrigen = adrwNovedad(ClsIdTipoDocOrigenByt.SstrNombreCampoBd)
            lenuTipoNovedad = adrwNovedad(ClsIdTipoNovedadByt.SstrNombreCampoBd)
        End If
        Dim lstrDetalle = FstrDetalleIni(lenuTipoNovedad)
        If lstrDetalle = String.Empty Then
            Select Case lenuTipoNovedad
                Case EnuTipoNov.EnuRCrAnApCap
                    lstrDetalle = "Anticipo aplicado a Capital Reversado"
                Case EnuTipoNov.EnuRCrAnApInt
                    lstrDetalle = "Anticipo aplicado a Int. de Mora Reversado"
                Case EnuTipoNov.EnuRCrDctoCap
                    lstrDetalle = "Descuento a Capital Reversado"
                Case EnuTipoNov.EnuRCrIvaGas
                    lstrDetalle = "Iva llevado al Gasto Reversado"
                Case EnuTipoNov.EnuRCrDctoInt
                    lstrDetalle = "Descuento a Int. de Mora Reversado"
                Case EnuTipoNov.EnuRCrPagoCap
                    lstrDetalle = "Pago a Capital Reversado"
                Case EnuTipoNov.EnuRCrPagoInt
                    lstrDetalle = "Pago a Int. de Mora Reversado"
                Case EnuTipoNov.EnuRCrRetFte
                    lstrDetalle = "Retención en la Fuente  Reversada"
                Case EnuTipoNov.EnuRCrRetIca
                    lstrDetalle = "Retención de Ind. y Comercio  Reversada"
                Case EnuTipoNov.EnuRCrRetIva
                    lstrDetalle = "Retención del Iva Reversada"
                Case EnuTipoNov.EnuRCrAntRec
                    lstrDetalle = "Anticipo recibido Reversado"
                Case EnuTipoNov.EnuRDbAntApl
                    lstrDetalle = "Anticipo aplicado Reversado"
                Case EnuTipoNov.EnuRDbAntDev
                    lstrDetalle = "Anticipo reintegrado Reversado"
            End Select
        End If
        Select Case lenuTipodocOrigen
            Case EnuTipoDocOri.EnuFactura
                lstrDetalle &= " Fra. Nro. "
            Case EnuTipoDocOri.EnuNotaDb
                lstrDetalle &= " Nota de Intereses Nro. "
            Case EnuTipoDocOri.EnuNotaCon
                lstrDetalle &= " Nota Aplicacion Anticipo Nro. "
            Case EnuTipoDocOri.EnuNotaCr
                lstrDetalle &= " Nota Crédito Nro. "
            Case EnuTipoDocOri.EnuNotaDevAnt
                lstrDetalle &= " Nota Reintegro Anticipo Nro. "
            Case EnuTipoDocOri.EnuReciboCaja
                lstrDetalle &= " Recibo de Caja Nro. "
            Case EnuTipoDocOri.EnuNotaRevCr
                lstrDetalle &= " Nota Reversión Recibo Caja "
        End Select
        lstrDetalle &= lstrDocOri
        Return lstrDetalle
    End Function
    Private Shared Function FstrDetalleIni(aenuTipoNovedad As EnuTipoNov) As String
        Dim lstrDetalle = String.Empty
        Select Case aenuTipoNovedad
            Case EnuTipoNov.EnuDbAntDev
                lstrDetalle = "Anticipo reintegrado"
            Case EnuTipoNov.EnuRDbCap
                lstrDetalle = "Capital Reversado"
            Case EnuTipoNov.EnuRDbInt
                lstrDetalle = "Intereses de Mora Reversados"
            Case EnuTipoNov.EnuRDbIva
                lstrDetalle = "IVA Reversado"
            Case EnuTipoNov.EnuCrIvaGas
                lstrDetalle = "IVA llevado al Gasto"
            Case EnuTipoNov.EnuDbCap
                lstrDetalle = "Valor Facturado"
            Case EnuTipoNov.EnuDbInt
                lstrDetalle = "Intereses de Mora"
            Case EnuTipoNov.EnuDbIva
                lstrDetalle = "IVA facturado"
            Case EnuTipoNov.EnuCrAnApCap
                lstrDetalle = "Anticipos aplicados a Capital"
            Case EnuTipoNov.EnuCrAnApInt
                lstrDetalle = "Anticipos aplicados a Int. de Mora"
            Case EnuTipoNov.EnuCrDctoCap
                lstrDetalle = "Descuento a Capital"
            Case EnuTipoNov.EnuCrDctoInt
                lstrDetalle = "Descuento a Int. de Mora"
            Case EnuTipoNov.EnuCrPagoCap
                lstrDetalle = "Pago a Capital"
            Case EnuTipoNov.EnuCrPagoInt
                lstrDetalle = "Pago a Int. de Mora"
            Case EnuTipoNov.EnuCrRetFte
                lstrDetalle = "Retención en la Fuente Aplicada"
            Case EnuTipoNov.EnuCrRetIca
                lstrDetalle = "Retención de Ind. y Comercio Aplicada"
            Case EnuTipoNov.EnuCrRetIva
                lstrDetalle = "Retención del Iva"
            Case EnuTipoNov.EnuCrAntRec
                lstrDetalle = "Anticipo recibido"
            Case EnuTipoNov.EnuDbAntApl
                lstrDetalle = "Anticipo aplicado"
            Case EnuTipoNov.EnuDbIvaInt
                lstrDetalle = "IVA a intereses de mora"
            Case EnuTipoNov.EnuRDbIvaInt
                lstrDetalle = "IVA a intereses de mora reversado"
        End Select
        Return lstrDetalle
    End Function
#End Region

#Region "Predios del Cliente"
    Friend Function FstrPrediosDelCliente() As String()
        Dim lstrPrediosDelCliente As String() = {}
        Dim lstrIdCliente = ObjIdClienteDbl.ToString()
        If String.IsNullOrEmpty(lstrIdCliente) Then lstrIdCliente = "0"
        If EnuEstadoActualizacion <> EnuEstadoObjetoDef.EnuCreando Then
            Dim lstrTabla = ClsPropietario.SstrNombreTabla
            Dim lstrCampSel As String() = {ClsIdPredio_PropStr.SstrNombreCampoBd}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsIdCliente_PropDbl.SstrNombreCampoBd & " = " & lstrIdCliente
            Dim lstrOrden As String(,) = {{ClsIdPredio_PropStr.SstrNombreCampoBd, "ASC"}}
            Dim ldtbPredios = ClsPanorama.FdtbDataTable(lstrTabla, lstrCampSel, lstrOrden,
                    lstrFiltro, False, {}), i = 0
            ReDim lstrPrediosDelCliente(ldtbPredios.Rows.Count - 1)
            For Each ldrwPredio As DataRow In ldtbPredios.Rows
                lstrPrediosDelCliente(i) = ClsPanorama.FobjValorCampo(
                    ldrwPredio(ClsIdPredio_PropStr.SstrNombreCampoBd),
                    EnuTipoValor.EnuString)
                i += 1
            Next
        End If
        Return lstrPrediosDelCliente
    End Function
    Private Function FstrPrediosAgruPropDelCliente() As String()
        Dim ldtbPreAgruCliente As DataTable = FdtbPrediosAgruPropDelCliente(), i = 0
        Dim lstrPrediosAgruDelCliente(ldtbPreAgruCliente.Rows.Count - 1) As String
        For Each ldrwPreAgr As DataRow In ldtbPreAgruCliente.Rows
            lstrPrediosAgruDelCliente(i) = ldrwPreAgr(ClsIdPredioAgrupadorStr.SstrNombreCampoBd)
            i += 1
        Next
        Return lstrPrediosAgruDelCliente
    End Function
    Private Function FstrPrediosAgruArrendados() As String()
        Dim ldtbPreAgruArrendados = FdtbPrediosArrendados(), i = 0
        Dim lstrPrediosArrendados(ldtbPreAgruArrendados.Rows.Count - 1) As String
        For Each ldrwPreAgr As DataRow In ldtbPreAgruArrendados.Rows
            lstrPrediosArrendados(i) = ldrwPreAgr(ClsIdPredioAgrupadorStr.SstrNombreCampoBd)
        Next
        Return lstrPrediosArrendados
    End Function
    Friend Function FstrPrediosAgruClienteConFacturas(ablnConSaldo As Boolean) As String()
        Dim lstrTabla = ClsFactura.SstrNombreTabla, i = 0
        Dim lstrCamposSelect As String() = {"DISTINCT " &
                ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdCliente_FactDbl.SstrNombreCampoBd & " = " & ObjIdClienteDbl.ToString
        If ablnConSaldo Then
            lstrFiltro &= " AND " & ClsDebitos_FactDec.SstrNombreCampoBd & " <> " &
                ClsCreditos_FactDec.SstrNombreCampoBd
        End If
        Dim lstrOrden As String(,) = {{ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd, "ASC"}}
        Dim ldtbPrediosAgr = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamposSelect, lstrOrden,
                lstrFiltro)
        Dim lstrPreAgr(ldtbPrediosAgr.Rows.Count - 1) As String
        For Each ldrwPreAgr As DataRow In ldtbPrediosAgr.Rows
            lstrPreAgr(i) = ldrwPreAgr(ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd)
            i += 1
        Next
        Array.Sort(lstrPreAgr)
        Return lstrPreAgr
    End Function
    Friend Function FstrPrediosAgruClienteTodos(ablnConFacturas As Boolean) As String()
        Dim lstrPreAgrPropCli = FstrPrediosAgruPropDelCliente()
        Dim lstrPregAgrArrendados = FstrPrediosAgruArrendados()
        Dim i = 0
        Dim lstrPreAgrCli As String() = {}
        For Each lstrPreAgr As String In lstrPreAgrPropCli
            If Not lstrPreAgrCli.Contains(lstrPreAgr) Then
                ReDim Preserve lstrPreAgrCli(i)
                lstrPreAgrCli(i) = lstrPreAgr
                i += 1
            End If
        Next
        For Each lstrPreAgr As String In lstrPregAgrArrendados
            If Not lstrPreAgrCli.Contains(lstrPreAgr) Then
                ReDim Preserve lstrPreAgrCli(i)
                lstrPreAgrCli(i) = lstrPreAgr
                i += 1
            End If
        Next
        If ablnConFacturas Then
            Dim lstrPreAgrConFac = FstrPrediosAgruClienteConFacturas(True)
            For Each lstrPreAgr As String In lstrPreAgrConFac
                If Not lstrPreAgrCli.Contains(lstrPreAgr) Then
                    ReDim Preserve lstrPreAgrCli(i)
                    lstrPreAgrCli(i) = lstrPreAgr
                    i += 1
                End If
            Next
        End If
        If lstrPreAgrCli.Length = 0 Then
            ReDim lstrPreAgrCli(0)
            lstrPreAgrCli(0) = ""
        End If
        Array.Sort(lstrPreAgrCli)
        Return lstrPreAgrCli
    End Function
    '
    Friend Function FdtbPrediosPropDelCliente() As DataTable
        Dim lstrTabla = ClsPropietario.SstrNombreTabla
        Dim lstrCamSel As String() = {ClsIdPredio_PropStr.SstrNombreCampoBd}
        Dim lstrOrden As String(,) = {{ClsIdPredio_PropStr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdCliente_PropDbl.SstrNombreCampoBd & " = " & ObjIdClienteDbl.ToString
        Dim ldtbPredCli = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamSel, lstrOrden, lstrFiltro)
        Return ldtbPredCli
    End Function
    Friend Function FdtbPrediosAgruPropDelCliente() As DataTable
        Dim lstrTablaPri = ClsPredio.SstrNombreTabla
        Dim lstrTablaSec = ClsPropietario.SstrNombreTabla
        Dim lstrCamSelPri As String() = {"DISTINCT " &
                ClsIdPredioAgrupadorStr.SstrNombreCampoBd}
        Dim lstrCamSelSec As String() = {}
        Dim lstrCamRelPri As String() = {StrCampoCarpeta,
                StrCampoCentroUtil, ClsIdPredioStr.SstrNombreCampoBd}
        Dim lstrCamRelSec As String() = {StrCampoCarpeta,
                StrCampoCentroUtil,
                ClsIdPredio_PropStr.SstrNombreCampoBd}
        Dim lstrOrden As String(,) = {{ClsIdPredioAgrupadorStr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion_Pri & " AND S." &
                ClsIdCliente_PropDbl.SstrNombreCampoBd & " = " & ObjIdClienteDbl.ToString
        Dim ldtbPredAgrCli = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCamSelPri,
                lstrTablaSec, lstrCamSelSec, lstrCamRelPri, lstrCamRelSec, lstrOrden,
                True, lstrFiltro, {})
        Return ldtbPredAgrCli
    End Function
    Friend Function FdtbPrediosArrendados() As DataTable
        Dim lstrTabla = ClsPredio.SstrNombreTabla
        Dim lstrCampSel As String() = {"DISTINCT " &
                ClsIdPredioAgrupadorStr.SstrNombreCampoBd}
        Dim lstrOrden As String(,) = {{ClsIdPredioAgrupadorStr.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdTipoDestinatarioFacturaByt.SstrNombreCampoBd & " = " &
                EnuDestinatarioFacturaDef.EnuArrendatario & " AND " &
                ClsIdClienteArrendatarioDbl.SstrNombreCampoBd & " = " &
                ObjIdClienteDbl.ToString
        Dim ldtbPredArrendados = ClsPanorama.FdtbDataTable(lstrTabla, lstrCampSel,
                lstrOrden, lstrFiltro)
        Return ldtbPredArrendados
    End Function
    Friend Function FblnPredioPropDelCliente(astrIdpredio As String) As Boolean
        Dim ldtbPreCli = FdtbPrediosPropDelCliente()
        Dim lstrFiltro = ClsIdPredio_PropStr.SstrNombreCampoBd & " = '" & astrIdpredio & "'"
        Return ldtbPreCli.Select(lstrFiltro).Length > 0
    End Function
    Friend Function FblnPredioEsArrendado(astrIdpredio As String) As Boolean
        Dim ldtbPreCli = FdtbPrediosArrendados()
        Dim lstrFiltro = ClsIdPredioAgrupadorStr.SstrNombreCampoBd & " = '" &
                astrIdpredio & "'"
        Return ldtbPreCli.Select(lstrFiltro).Length > 0
    End Function
#End Region

#Region "Recibos de Caja"
    Friend Function FdtbRecibos(astrIdPredioAgr As String, adtmFechaDesde As Date,
             adtmFechaHasta As Date) As DataTable
        SCargueDtbRecibos(astrIdPredioAgr, adtmFechaDesde, adtmFechaHasta)
        Return MdtbRecibos
    End Function
    Private Sub SCargueDtbRecibos(astrIdPredioAgru As String, ByRef adtmFechaDesde As Date,
                                    adtmFechaHasta As Date)
        Dim lstrExpSqlRecCajaNov = FstrExpSqlRecCajaNov(astrIdPredioAgru, adtmFechaDesde,
                adtmFechaHasta)
        Dim lstrExpSqlRecCajNovAnt = FstrExpSqlRecCajaNovAnt(astrIdPredioAgru, adtmFechaDesde,
                adtmFechaHasta)
        Dim lstrIndice = ClsFechaRecDtm.SstrNombreCampoBd & ", " &
                  ClsPrefijo_RecStr.SstrNombreCampoBd & ", " &
                  ClsIdRecCajaEnt.SstrNombreCampoBd
        Dim lstrExpSql = "(" & lstrExpSqlRecCajaNov & ") UNION ALL (" & lstrExpSqlRecCajNovAnt +
                ") ORDER BY " & lstrIndice
        GobjPanDat.SControleProcesoObj(True)
        MdtbRecibos = ClsPanorama.FdtbDataTable(lstrExpSql)
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Private Function FstrExpSqlRecCajaNov(astrIdPredioAgru As String, ByRef adtmFechaDesde As Date,
             adtmFechaHasta As Date) As String
        Dim lstrFechaDesde = ClsPanoramaDat.FstrFechaNormalizada(adtmFechaDesde)
        Dim lstrFechaHasta = ClsPanoramaDat.FstrFechaNormalizada(adtmFechaHasta)
        Dim lstrTablaPri = ClsNovedad.SstrNombreTabla
        Dim lstrTablaSec = ClsReciboCaja.SstrNombreTabla
        Dim lstrCamposPri() = {"DISTINCT " & ClsFechaNovedadDtm.SstrNombreCampoBd}
        Dim lstrCamposSec() = {"*"}
        Dim lstrCamposRelPri = {StrCampoCarpeta,
                                StrCampoCentroUtil,
                                ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd,
                                ClsIdDocOrigenEnt.SstrNombreCampoBd}
        Dim lstrCamposRelSec = {StrCampoCarpeta,
                                StrCampoCentroUtil,
                                ClsPrefijo_RecStr.SstrNombreCampoBd,
                                ClsIdRecCajaEnt.SstrNombreCampoBd}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion_Pri & " AND " &
                ClsIdTipoDocOrigenByt.SstrNombreCampoBd & " = 2 AND S." &
                ClsIdCliente_RecDbl.SstrNombreCampoBd & " = " & ObjIdClienteDbl.ToString &
                " AND P." & ClsIdPredioAgrupador_NovStr.SstrNombreCampoBd & " = '" &
                astrIdPredioAgru & "'"
        lstrFiltro += " AND " & ClsFechaRecDtm.SstrNombreCampoBd & " >= '" &
                lstrFechaDesde + "' AND " & ClsFechaRecDtm.SstrNombreCampoBd & " <= '" &
                lstrFechaHasta & "'"
        Dim lstrSql As String = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTablaPri,
                lstrCamposPri, lstrTablaSec, lstrCamposSec, lstrCamposRelPri,
                lstrCamposRelSec, {{"", ""}}, lstrFiltro, Array.Empty(Of String))
        Return lstrSql
    End Function
    Private Function FstrExpSqlRecCajaNovAnt(astrIdPredioAgru As String, ByRef adtmFechaDesde As Date,
             adtmFechaHasta As Date) As String
        Dim lstrFechaDesde = ClsPanoramaDat.FstrFechaNormalizada(adtmFechaDesde)
        Dim lstrFechaHasta = ClsPanoramaDat.FstrFechaNormalizada(adtmFechaHasta)
        Dim lstrTablaPri = ClsNovedadAnticipo.SstrNombreTabla
        Dim lstrTablaSec = ClsReciboCaja.SstrNombreTabla
        Dim lstrCamposPri() = {"DISTINCT " & ClsFechaNovedadAntDtm.SstrNombreCampoBd}
        Dim lstrCamposSec() = {"*"}
        Dim lstrCamposRelPri = {StrCampoCarpeta,
                                StrCampoCentroUtil,
                                ClsPrefijoDocOrigen_NovAntStr.SstrNombreCampoBd,
                                ClsIdDocOrigen_NovAntEnt.SstrNombreCampoBd}
        Dim lstrCamposRelSec = {StrCampoCarpeta,
                                StrCampoCentroUtil,
                                ClsPrefijo_RecStr.SstrNombreCampoBd,
                                ClsIdRecCajaEnt.SstrNombreCampoBd}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion_Pri & " AND " &
                ClsIdTipoDocOrigen_NovAntByt.SstrNombreCampoBd & " = 2 AND S." &
                ClsIdCliente_RecDbl.SstrNombreCampoBd & " = " & ObjIdClienteDbl.ToString &
                " AND S." & ClsIdPredioAgrupador_RecStr.SstrNombreCampoBd & " = '" &
                astrIdPredioAgru & "' AND S." & ClsValor_RecDec.SstrNombreCampoBd &
                " = " & ClsValorAnticipoDec.SstrNombreCampoBd
        lstrFiltro &= " AND " & ClsFechaRecDtm.SstrNombreCampoBd & " >= '" &
                lstrFechaDesde & "' AND " & ClsFechaRecDtm.SstrNombreCampoBd & " <= '" &
                lstrFechaHasta & "'"
        Dim lstrSql As String = ClsPanoramaDat.FstrConstruyaExpSqlSelect(lstrTablaPri,
                lstrCamposPri, lstrTablaSec, lstrCamposSec, lstrCamposRelPri,
                lstrCamposRelSec, {{"", ""}}, lstrFiltro, Array.Empty(Of String))
        Return lstrSql
    End Function
#End Region

#Region "Retenciones y Descuentos"
    ''' <summary>
    ''' Devuelve un DataTable con la información de los descuentos que se deben hacer a cada 
    ''' una de las facturas con saldo pertenecientes a los Predios Agrupadores indicados 
    ''' en el parametro "astrIdPredioAgru" y al Agrupador de Servicios indicado en el 
    ''' parametro "ashrIdAgruServicios"
    ''' </summary>
    Friend Function FdtbDescuentos(astrIdPrediosAgru As String(), astrServicios As String(),
            adtmFechaPago As Date) As DataTable
        Dim ldtbDescuentos As DataTable = ClsOrionCop.FdtbDescuentos()
        If ObjEsAgenteReteFteBln.ObjValorPro OrElse ObjRetieneIcaBln.ObjValorPro OrElse
                ObjRetieneIvaBln.ObjValorPro OrElse
                GobjParametros.ObjAnoActual.ObjTipoIncentivoByt.ObjValorPro =
                    EnuTipoIncentivo.EnuDescuentoPP Then
            Dim lcolFras = FcolFacturas(astrIdPrediosAgru, astrServicios, True)
            For Each lobjFac As ClsFactura In lcolFras
                SPuebleDtbDescuentos(lobjFac, adtmFechaPago, astrServicios, ldtbDescuentos)
            Next
        End If
        Return ldtbDescuentos
    End Function
    Private Sub SPuebleDtbDescuentos(aobjFactura As ClsFactura, adtmFechaPago As Date,
            astrServicios As String(), adtbDsctos As DataTable)
        For Each lobjItemFac As ClsItemFactura In aobjFactura.ColItemsFactura
            If lobjItemFac.ObjIdAno_ServicioItemFactShr.ObjValorPro > 0 Then
                If astrServicios.Contains("A") OrElse astrServicios.Contains("0") Then
                    SPuebleDtbDescuentos(lobjItemFac, adtmFechaPago, adtbDsctos)
                End If
            Else
                SPuebleDtbDescuentos(lobjItemFac, adtmFechaPago, adtbDsctos)
            End If
        Next
    End Sub
    Private Sub SPuebleDtbDescuentos(aobjItemFactura As ClsItemFactura, adtmFechaPago As Date,
             ByRef adtbDsctos As DataTable)
        Dim ldecVlrRetencion As Decimal, lstrNombreDcto As String
        Dim ldecBaseDscto As Decimal = 0
        Dim ldblTasaDscto As Double = 0
        Dim ldecSaldoItemFac As Decimal = aobjItemFactura.DecDeuda
        Dim ldecDsctoProntoPago = aobjItemFactura.FdecDsctoPPAAplicar(adtmFechaPago)
        Dim ldecTotalDscto = 0D
        If ldecDsctoProntoPago > 0 AndAlso aobjItemFactura.FdecDeudaCapital > ldecTotalDscto Then
            ldecBaseDscto = aobjItemFactura.ObjValor_ItemFactDec.ObjValorPro
            ldblTasaDscto = ldecDsctoProntoPago / ldecBaseDscto
            ldecTotalDscto += ldecDsctoProntoPago
            Dim ldrwNuevoDscto = adtbDsctos.NewRow
            ldrwNuevoDscto("Ordinal") = adtbDsctos.Rows.Count + 1
            ldrwNuevoDscto("NroFactura") = aobjItemFactura.StrNumeroFactura
            ldrwNuevoDscto("IdItemFact") = aobjItemFactura.ObjIdItemFacturaShr.ToString & "-" &
                    aobjItemFactura.ObjDetalle_ItemFactStr.ObjValorPro
            ldrwNuevoDscto("IdTipoDcto") = EnuTipoDescuento.EnuDsctoPP
            ldrwNuevoDscto("Base") = ldecBaseDscto
            ldrwNuevoDscto("Tasa") = ldblTasaDscto
            ldrwNuevoDscto("Valor") = ldecDsctoProntoPago
            lstrNombreDcto = ClsOrionCop.FstrNombreDatoConstanteOri(
                    EnuGrupoConstantesOriDef.EnuTipoDescuento, EnuTipoDescuento.EnuDsctoPP)
            ldrwNuevoDscto("TipoDcto") = lstrNombreDcto
            adtbDsctos.Rows.Add(ldrwNuevoDscto)
        End If
        For i = EnuTipoDescuento.EnuReteFuente To EnuTipoDescuento.EnuReteCree
            ldecVlrRetencion = aobjItemFactura.FdecValorRetencion(i, ldecBaseDscto, ldblTasaDscto)
            If ldecVlrRetencion > 0 Then
                ldecTotalDscto += ldecVlrRetencion
                Dim ldrwNuevaReten = adtbDsctos.NewRow
                ldrwNuevaReten("Ordinal") = adtbDsctos.Rows.Count + 1
                ldrwNuevaReten("NroFactura") = aobjItemFactura.StrNumeroFactura
                ldrwNuevaReten("IdItemFact") = aobjItemFactura.ObjIdItemFacturaShr.ToString & "-" &
                        aobjItemFactura.ObjDetalle_ItemFactStr.ObjValorPro
                ldrwNuevaReten("IdTipoDcto") = i
                ldrwNuevaReten("Base") = ldecBaseDscto
                ldrwNuevaReten("Tasa") = ldblTasaDscto
                ldrwNuevaReten("Valor") = ldecVlrRetencion
                lstrNombreDcto = ClsOrionCop.FstrNombreDatoConstanteOri(
                        EnuGrupoConstantesOriDef.EnuTipoDescuento, i)
                ldrwNuevaReten("TipoDcto") = lstrNombreDcto
                adtbDsctos.Rows.Add(ldrwNuevaReten)
            End If
        Next
    End Sub
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsComentario_ClienteStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Comentario"
        HshrLongitud = 500
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "Comentario"
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 3, ShrLongitud, BlnEsRequerido)
    End Sub
    Public Overrides Function ToString() As String
        If Not IsNothing(HobjValorPro) Then
            Return HobjValorPro
        Else
            Return ""
        End If
    End Function
End Class
Friend Class ClsEsAgenteReteFteBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "EsAgenteRetenedor"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Es Agente Retenedor en la Fuente"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
        If GobjParametros.BlnEFacAutorizado Then
            HblnEsRequerido = True
        End If
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
Friend Class ClsEsAutorretenedorBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "EsAutoRetenedor"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Es autorretenedor"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
        If GobjParametros.BlnEFacAutorizado Then
            HblnEsRequerido = True
        End If
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
Friend Class ClsEsGranContrBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "EsGranContribuyente"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Es gran Contribuyente"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
        If GobjParametros.BlnEFacAutorizado Then
            HblnEsRequerido = True
        End If
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
Friend Class ClsEsRegimenSimpleTBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "EsRegimenSimpleTributacion"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Es Régimen Simple"
        HenuTipoValor = EnuTipoValor.enuBoolean
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
Friend Class ClsFactPorServicio_CliBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FacturarPorServicio"
    Private ReadOnly MobjPadre As ClsCliente = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Facturar por Agrupador de Servicios"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                If HobjValorNew <> HobjValorOriginal Then
                    HblnEsValido = MobjPadre.FdecDeuda({""}, {"A"}) = 0
                    If Not HblnEsValido Then
                        HstrMens = "Este cambio no es posible mientras el Cliente tenga Deuda!"
                        SNotifiqueDatInv()
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
        Return ClsPanorama.FstrBuleanoToString(HobjValorPro)
    End Function
End Class
Friend Class ClsIdEstadoDeudaByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdEstadoDeuda"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdEstadoDeuda"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuEstadoDeudaDef.EnuNormal,
                    EnuEstadoDeudaDef.EnuSuspendida, HblnEsRequerido)
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
            Return ClsOrionCop.FstrNombreDatoConstanteOri(EnuGrupoConstantesOriDef.EnuEstadoDeuda, HobjValorPro)
        End If
    End Function
End Class
Friend Class ClsIdMedioPagoClienteByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTipoMediosPago"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdTipoMediosPago"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = EnuTipoMedioPagoDef.None
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew,
                EnuTipoMedioPagoDef.EnuEfectivo, EnuTipoMedioPagoDef.EnuTransferencia, HblnEsRequerido)
    End Sub
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
Friend Class ClsIdRegimenVentasByt
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdRegimenVentas"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = "IdRegimenVentas"
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuRegimenVentasDef.EnuNoResponsable,
                    EnuRegimenVentasDef.EnuResponsable, HblnEsRequerido)
    End Sub
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsIdClienteDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTerceroCliente"
    Private ReadOnly MobjPadre As ClsCliente = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdCliente"
        MobjPadre = aobjPadre
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 2
    End Sub
    Public Overrides Sub SValide()
        HstrMens = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCDBLMINTERC,
                GCDBLMAXTERC, BlnEsRequerido)
        If HblnEsValido Then
            HblnEsValido = HobjValorNew <> GCDBLTERCERONULO
            If Not HblnEsValido Then
                HstrMens = "Este número de identificación está reservado por el programa!"
            End If
        End If
        If HblnEsValido Then
            If Not BlnLeyendoOrigen Then
                If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
                    Dim lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
                    If MobjPadre.FblnExisteLlave(lobjValorLlave) Then
                        HstrMens = "Un Cliente con el número de identificación ingresado, '" &
                                HobjValorNew.ToString & "', ya existe!"
                        HblnEsValido = False
                    End If
                    If HblnEsValido Then
                        Dim lblnExiste = MobjPadre.ObjTerceroCliente.FblnExisteLlave({HobjValorNew})
                        If MobjPadre.ObjTerceroCliente.EnuEstadoActualizacion =
                                EnuEstadoObjetoDef.EnuConsultando Then
                            If lblnExiste Then
                                HblnEsValido = True
                            Else
                                MobjPadre.ObjTerceroCliente.SCreeObj({HobjValorNew})
                                MobjPadre.ObjIdTerceroDbl.ObjValorPro = HobjValorNew
                            End If
                        Else
                            If Not lblnExiste Then
                                MobjPadre.ObjIdTerceroDbl.ObjValorPro = HobjValorNew
                            End If
                        End If
                    End If
                ElseIf MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando Then
                    HblnEsValido = (HobjValorOriginal = HobjValorNew)
                    If Not HblnEsValido Then
                        HstrMens = "No es permitido cambiar la identidad a objeto alguno!"
                    End If
                End If
            End If
        Else
            If (MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando AndAlso
                        HobjValorNew <> 0 AndAlso HobjValorNew <> GCDBLTERCERONULO) OrElse
                        MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
                HstrMens = "El valor ingresado, '" & HobjValorNew.ToString & "', no es valido!"
            End If
        End If
        If Not String.IsNullOrEmpty(HstrMens) Then
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
Friend Class ClsNombreCompletoStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "NombreCompleto"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "NombreCompleto"
        HshrLongitud = 100
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
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
Friend Class ClsRecibeDocsPorEmailBln
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsCliente = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "RecibeDocsPorEmail"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Recibe Emails"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        If GobjParametros.BlnEFacAutorizado Then
            HblnEsRequerido = True
        End If
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoBuleano(HobjValorNew)
        If HblnEsValido AndAlso HobjValorNew Then
            HblnEsValido = FblnTieneEmail()
            If Not HblnEsValido Then
                HstrMens = "El Cliente no tiene una Dirección de Correo Electrónico!"
                SNotifiqueDatInv()
            End If
        End If
    End Sub
    Private Sub EPosCambioVlr(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosCambio
        MobjPadre.ObjEmailStr.SValide()
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Private Function FblnTieneEmail() As Boolean
        Dim lblnTiene = False
        If MobjPadre.ObjIdClienteDbl.BlnEsValido Then
            If MobjPadre.ObjTerceroCliente IsNot Nothing Then
                Dim lstrEmail = MobjPadre.ObjTerceroCliente.ObjEmailStr.ObjValorPro
                lblnTiene = Not String.IsNullOrEmpty(lstrEmail)
            End If
        End If
        Return lblnTiene
    End Function
    Public Overrides Function ToString() As String
        Return ClsPanorama.FstrBuleanoToString(HobjValorPro)
    End Function
End Class
Friend Class ClsRetieneIcaBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "RetieneIca"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Reriene Inductria y Comercio"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
        If GobjParametros.BlnEFacAutorizado Then
            HblnEsRequerido = True
        End If
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
Friend Class ClsRetieneIvaBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "RetieneIva"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Retiene IVA"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
        If GobjParametros.BlnEFacAutorizado Then
            HblnEsRequerido = True
        End If
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
#End Region
Friend Class ClsPredio
#Region "Definiciones"
    ' Herencia 
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriPredios"
    ' 
    Private MdtbPrediosAgrupados As DataTable = Nothing
    Private McolPrediosAgrupados As Collection = Nothing
    Private McolPropietarios As Collection = Nothing
    Private MstrPrediosAgrupados() As String = Nothing
    Private MdtbItemsProgFact As DataTable = Nothing
    Private McolItemsProgramaFact As Collection = Nothing
    Private MdtbItemsFactura As DataTable = Nothing
    Private McolItemsFactura As Collection = Nothing
    Private MobjCliente As ClsCliente = Nothing
    Private MobjPropietario As ClsCliente = Nothing
    Private MobjArrendatario As ClsCliente = Nothing
    Private MobjSector As ClsSector = Nothing
    Private MshrIdAno As Short = -1
    Private MblnSoloVivasColeccion As Boolean = False
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia un objeto Predio.
    ''' </summary>
    ''' <param name="aenuModoInstanciaObj">Indica si se instancia como un objeto navegable o como un Objeto único.</param>
    ''' <remarks>Si se instancia como un objeto navegable, se crea un datatable que contiene las columnas de
    ''' la llave con las llaves de todos los objetos y queda a la espera de que se indique que objeto abrir.
    ''' Si se instancia como un objeto único, queda a la espera de recibir el valor de los campos de la llave 
    ''' para abrir dicho objeto. </remarks>
    Public Sub New(aenuModoInstanciaObj As EnuModoInstanciaObjDef)
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.EnuDeColeccion Then
            Throw New ErrorInesperadoPanLException("Con este Constructor no se puede instanciar un Objeto de Colección!")
        End If
        Dim lstrCamposSelect As String()
        HobjPadre = Nothing
        If aenuModoInstanciaObj = EnuModoInstanciaObjDef.EnuNavegable Then
            HblnEsAnulable = False
            lstrCamposSelect = {StrCampoCarpeta, StrCampoCentroUtil,
                    ClsIdPredioStr.SstrNombreCampoBd}
            HcolFiltros.Add(ClsOrionCop.StrFiltroUbicacion)
        Else
            HblnEsCreable = False
            HblnEsModificable = False
            HblnEsSuprimible = False
            HblnEsAnulable = False
            HenuTipoObjeto = EnuModoInstanciaObjDef.EnuUnico
            lstrCamposSelect = {"*"}
        End If
        HcolTablas.Add(MCSTRNOMBRETABLA)
        HcolCamposSelect.Add(lstrCamposSelect)
    End Sub
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="adrwObjeto">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Friend Sub New(adrwObjeto As DataRow)
        HobjPadre = Nothing
        HenuTipoObjeto = EnuModoInstanciaObjDef.EnuDeColeccion
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
            Return EnuIdClasesPanDef.EnuPredio
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "El Predio"
        End Get
    End Property
    Friend Overrides ReadOnly Property HstrNombreObj As String
        Get
            Return Chr(34) & ObjIdPredioStr.ObjValorPro & Chr(34)
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjAliasContStr As New ClsAliasContStr(Me)
    Friend ReadOnly Property ObjAreaPredioDec As New ClsAreaPredioDec(Me)
    Friend ReadOnly Property ObjCoeficientePropiedadDec As New ClsCoeficientePropiedadDec(Me)
    Friend ReadOnly Property ObjComentarioStr As New ClsComentarioStr(Me)
    Friend ReadOnly Property ObjEmailAdiStr As New ClsEmailAdiStr(Me)
    Friend ReadOnly Property ObjFactorPonderaCPDbl As New ClsFactorPonderaCPDbl(Me)
    Friend ReadOnly Property ObjFacturarPorServicio_PreBln As New ClsFacturarPorServicio_PreBln(Me)
    Friend ReadOnly Property ObjIdCarpeta_PredioShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_PredioShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdEstadoDeuda_PredioByt As New ClsIdEstadoDeuda_PredioByt(Me)
    Friend ReadOnly Property ObjIdFichaCatastralStr As New ClsIdFichaCatastralStr(Me)
    Friend ReadOnly Property ObjIdMatriculaInmobiliariaStr As New ClsIdMatriculaInmobiliariaStr(Me)
    Friend ReadOnly Property ObjIdPredioAgrupadorStr As New ClsIdPredioAgrupadorStr(Me)
    Friend ReadOnly Property ObjIdPredioStr As New ClsIdPredioStr(Me)
    Friend ReadOnly Property ObjIdRegistroMercantilStr As New ClsIdRegistroMercantilStr(Me)
    Friend ReadOnly Property ObjIdSector_PredioShr As New ClsIdSector_PredioShr(Me)
    Friend ReadOnly Property ObjIdClienteAdministradorDbl As New ClsIdClienteAdministradorDbl(Me)
    Friend ReadOnly Property ObjIdClienteArrendatarioDbl As New ClsIdClienteArrendatarioDbl(Me)
    Friend ReadOnly Property ObjIdClienteRepLegArrendatariodbl As New ClsIdClienteRepLegArrendatariodbl(Me)
    Friend ReadOnly Property ObjIdTipoDestinatarioFacturaByt As New ClsIdTipoDestinatarioFacturaByt(Me)
    Friend ReadOnly Property ObjNoConsolidarItemsFacBln As New ClsNoConsolidarItemsFacBln(Me)
    Friend ReadOnly Property ObjNombreComercialStr As New ClsNombreComercialStr(Me)
    Friend ReadOnly Property ObjNombrePredioStr As New ClsNombrePredioStr(Me)
    Friend ReadOnly Property ObjReferenciaPagoStr As New ClsReferenciaPagoStr(Me)
    Friend ReadOnly Property ObjValorServicioIdDec As New ClsValorServicioIdDec(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjAliasContStr)
                HcolPropiedades.Add(ObjAreaPredioDec)
                HcolPropiedades.Add(ObjCoeficientePropiedadDec)
                HcolPropiedades.Add(ObjComentarioStr)
                HcolPropiedades.Add(ObjEmailAdiStr)
                HcolPropiedades.Add(ObjFactorPonderaCPDbl)
                HcolPropiedades.Add(ObjIdCarpeta_PredioShr)
                HcolPropiedades.Add(ObjIdCentroUtil_PredioShr)
                HcolPropiedades.Add(ObjIdEstadoDeuda_PredioByt)
                HcolPropiedades.Add(ObjIdFichaCatastralStr)
                HcolPropiedades.Add(ObjIdMatriculaInmobiliariaStr)
                HcolPropiedades.Add(ObjIdPredioStr)
                HcolPropiedades.Add(ObjIdPredioAgrupadorStr)
                HcolPropiedades.Add(ObjIdRegistroMercantilStr)
                HcolPropiedades.Add(ObjIdSector_PredioShr)
                HcolPropiedades.Add(ObjIdClienteAdministradorDbl)
                HcolPropiedades.Add(ObjIdClienteArrendatarioDbl)
                HcolPropiedades.Add(ObjIdClienteRepLegArrendatariodbl)
                HcolPropiedades.Add(ObjIdTipoDestinatarioFacturaByt)
                HcolPropiedades.Add(ObjNoConsolidarItemsFacBln)
                HcolPropiedades.Add(ObjNombreComercialStr)
                HcolPropiedades.Add(ObjNombrePredioStr)
                HcolPropiedades.Add(ObjReferenciaPagoStr)
                HcolPropiedades.Add(ObjValorServicioIdDec)
                HcolPropiedades.Add(ObjFacturarPorServicio_PreBln)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras Propiedades"
    ''' <summary>
    ''' Devuelve un arreglo que contiene la identificación de los predios agrupados, incluido el predio
    ''' agrupador.
    ''' </summary>
    Friend ReadOnly Property StrPrediosAgrupados As String()
        Get
            If IsNothing(MstrPrediosAgrupados) Then
                Try
                    SPueblePrediosAgrupados()
                Catch ex As ErrorInesperadoPanLException
                    Throw
                End Try
            End If
            Return MstrPrediosAgrupados
        End Get
    End Property
    Friend ReadOnly Property ObjPredioAgrupador As ClsPredio
        Get
            Dim lobjPredioAgr As ClsPredio
            If ObjIdPredioStr.ObjValorPro = ObjIdPredioAgrupadorStr.ObjValorPro Then
                lobjPredioAgr = Me
            Else
                Dim lobjValorLlave() As Object = {GshrIdCarpeta, GshrIdCentroUtil,
                                                  ObjIdPredioAgrupadorStr.ObjValorPro}
                lobjPredioAgr = New ClsPredio(EnuModoInstanciaObjDef.EnuUnico)
                lobjPredioAgr.SAbra(lobjValorLlave)
            End If
            Return lobjPredioAgr
        End Get
    End Property
    Friend ReadOnly Property StrNombrePredioAgrupador As String
        Get
            Return ObjPredioAgrupador.ObjNombrePredioStr.ObjValorPro
        End Get
    End Property
    Friend ReadOnly Property ObjArrendatario As ClsCliente
        Get
            If ObjIdClienteArrendatarioDbl.BlnEsValido AndAlso
                    ObjIdClienteArrendatarioDbl.ObjValorPro > 0 Then
                Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil,
                              ObjIdClienteArrendatarioDbl.ObjValorPro}
                MobjArrendatario = New ClsCliente(EnuModoInstanciaObjDef.EnuUnico)
                MobjArrendatario.SAbra(lobjValorLlave)
            End If
            Return MobjArrendatario
        End Get
    End Property
    ''' <summary>
    ''' Devuelve una colección con los predios agrupados bajo este predio incluyendo el predio agrupador.
    ''' </summary>
    Friend ReadOnly Property ColPrediosAgrupados As Collection
        Get
            If IsNothing(McolPrediosAgrupados) Then
                McolPrediosAgrupados = New Collection
                If ObjIdPredioAgrupadorStr.ObjValorPro = ObjIdPredioStr.ObjValorPro Then
                    SPuebleTablaPrediosAgrupados()
                    If MdtbPrediosAgrupados.Rows.Count > 0 Then
                        Dim ldrwPrediosAgrupados As DataRow() = MdtbPrediosAgrupados.Select()
                        For Each ldrwPredioAgr As DataRow In ldrwPrediosAgrupados
                            Dim lobjPredio As New ClsPredio(ldrwPredioAgr)
                            lobjPredio.SLeaValores(True)
                            McolPrediosAgrupados.Add(lobjPredio, lobjPredio.ObjIdPredioStr.ToString)
                        Next
                    End If
                End If
            End If
            Return McolPrediosAgrupados
        End Get
    End Property
    Friend ReadOnly Property ObjSector As ClsSector
        Get
            If IsNothing(MobjSector) Then
                MobjSector = GobjParametros.ColSectores(ObjIdSector_PredioShr.ToString)
            End If
            Return MobjSector
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SVacie()
        MyBase.SVacie()
        MstrPrediosAgrupados = Nothing
        MdtbItemsProgFact = Nothing
        McolItemsProgramaFact = Nothing
        McolPrediosAgrupados = Nothing
        McolPropietarios = Nothing
        MdtbPrediosAgrupados = Nothing
        MdtbItemsFactura = Nothing
        McolItemsFactura = Nothing
        MobjCliente = Nothing
        MobjPropietario = Nothing
        MobjArrendatario = Nothing
        MobjSector = Nothing
        MshrIdAno = -1
    End Sub
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        GobjPanDat.SControleProcesoObj(True)
        Dim lblnNoHayError = False, lblnPropValidos = True
        Try
            If EnuEstadoActualizacion <> EnuEstadoObjetoDef.EnuConsultando Then
                GobjPanDat.SInicialiceTransaccion()
                Dim lblnCambioPro As Boolean
                SActualicePropietarios(lblnCambioPro)
                If Not lblnCambioPro Then
                    lblnCambioPro = FblnCambioPropietario()
                End If
                ClsPanorama.SActualiceCol(ColPropietarios)
                If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando Then
                    If ObjIdPredioAgrupadorStr.BlnCambio AndAlso
                        ObjIdPredioStr.ObjValorPro =
                        ObjIdPredioAgrupadorStr.ObjValorOriginal Then
                        SActuaPredAgruAPrediosAgrupados()
                    End If
                    If lblnCambioPro AndAlso ColPrediosAgrupados.Count > 0 Then
                        SActualicePropPredios()
                    End If
                    SActualicePrediosAgrupados()
                    MyBase.SActualice(ablnExigeRequeridos)
                Else
                    lblnPropValidos = FblnPropsValidos()
                    If lblnPropValidos Then
                        MyBase.SActualice(ablnExigeRequeridos)
                    End If
                End If
            End If
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
                If lblnPropValidos Then
                    GobjPanDat.SConfirmeTransaccion()
                    GobjPanDat.SControleProcesoObj(False)
                Else
                    GobjPanDat.SAborteTransaccion()
                    GobjPanDat.SControleProcesoObj(False, True)
                    SLevanteEventoNot("No hay integridad en los propietarios! " &
                            "El Predio " & ObjIdPredioStr.ToString & " no fue creado!",
                            String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                End If
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    Protected Overrides Sub SInicialiceObj()
        MyBase.SInicialiceObj()
        ObjIdCarpeta_PredioShr.ObjValorPro = GshrIdCarpeta
        ObjIdCentroUtil_PredioShr.ObjValorPro = GshrIdCentroUtil
        ObjAreaPredioDec.ObjValorPro = 0
        ObjFactorPonderaCPDbl.ObjValorPro = 0
        ObjNoConsolidarItemsFacBln.ObjValorPro = True
        ObjFacturarPorServicio_PreBln.ObjValorPro = False
    End Sub
    Protected Overrides Function FblnSuprimio() As Boolean
        Dim lblnSuprimio = FblnEsSuprimible()
        If lblnSuprimio Then
            lblnSuprimio = MyBase.FblnSuprimio()
            If lblnSuprimio AndAlso BlnEsNavegable Then
                GobjPanorama.SRegistreAccionLogApp(HstrNombreClase, "Suprimir Predio con Id. " &
                    ObjIdPredioStr.ObjValorPro)
            End If
        End If
        Return lblnSuprimio
    End Function
    Friend Overrides Function FblnEsSuprimible() As Boolean
        Dim lblnEsSuprimible = FblnPermitidoSuprimir()
        If lblnEsSuprimible Then
            SCargueDtbItemsFactura()
            lblnEsSuprimible = (MdtbItemsFactura.Rows.Count = 0)
            If lblnEsSuprimible Then
                Dim lstrCondicion As String = " = '" & ObjIdPredioStr.ObjValorPro & "' AND " &
                    ClsOrionCop.StrFiltroUbicacion
                lblnEsSuprimible = ClsPanorama.FblnEsEliminableReg({SstrNombreTabla},
                    ObjIdPredioStr.StrNombreCampoBD, lstrCondicion, True, False)
            End If
        End If
        Return lblnEsSuprimible
    End Function
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjIdPredioStr.ToString
        End Get
    End Property
    Friend Overrides Function FblnSonValidosDatosOrigen(adtbOrigen As DataTable,
            astrColumnasRelacionadas As String(), ablnReinicie As Boolean,
            ByRef astrMens As String) As Boolean
        Dim lblnEsValido = False, i = 0, ldblFactPondCP As Double, lstrRefPago As String,
                lstrColumnaOrigen As String
        Dim lbytIdCar As Byte, lbytIdCenutil As Byte, lbytIdSector As Byte,
                lblnConRefPago As Boolean
        Dim lobjSector As New ClsSector(EnuModoInstanciaObjDef.EnuUnico)
        Dim lobjValorLlave As Object()
        For Each ldrwOrigen As DataRow In adtbOrigen.Rows
            i += 1
            lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                    ClsIdCarpetaShr.SstrNombreCampoBd)
            lbytIdCar = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrColumnaOrigen),
                    EnuTipoValor.EnuByte)
            lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                    ClsIdCentroUtilShr.SstrNombreCampoBd)
            lbytIdCenutil = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrColumnaOrigen),
                    EnuTipoValor.EnuByte)
            lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                    ClsFactorPonderaCPDbl.SstrNombreCampoBd)
            ldblFactPondCP = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrColumnaOrigen),
                    EnuTipoValor.EnuDouble)
            lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                    ClsIdSector_PredioShr.SstrNombreCampoBd)
            lbytIdSector = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrColumnaOrigen),
                    EnuTipoValor.EnuByte)
            lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                    ClsReferenciaPagoStr.SstrNombreCampoBd)
            lstrRefPago = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrColumnaOrigen),
                    EnuTipoValor.EnuString)
            lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lbytIdSector}
            lobjSector.SAbra(lobjValorLlave)
            lblnEsValido = lobjSector.BlnExiste
            If Not lblnEsValido Then
                astrMens = "El Sector correspondiente al Predio del registro " & i.ToString &
                        " no existe!"
                Exit For
            End If
            lblnEsValido = lbytIdCar = GshrIdCarpeta
            If Not lblnEsValido Then
                astrMens = "La Carpeta del registro " & i.ToString & " no es la Carpeta actual!"
                Exit For
            End If
            lblnEsValido = lbytIdCenutil = GshrIdCentroUtil
            If Not lblnEsValido Then
                astrMens = "La Copropiedad del registro " & i.ToString & " no es el actual!"
                Exit For
            End If
            lblnEsValido = ldblFactPondCP >= 0 AndAlso ldblFactPondCP <= 1
            If Not lblnEsValido Then
                astrMens = "El factor de ponderación del registro " & i.ToString & " no es valido!"
                Exit For
            End If
            If i = 1 Then
                lblnConRefPago = Not String.IsNullOrEmpty(lstrRefPago)
            Else
                If lblnConRefPago Then
                    lblnEsValido = Not String.IsNullOrEmpty(lstrRefPago)
                Else
                    lblnEsValido = String.IsNullOrEmpty(lstrRefPago)
                End If
                If Not lblnEsValido Then
                    astrMens = "La Referencia de Pago en el registro" & i.ToString & " no es valida"
                    Exit For
                End If
            End If
        Next
        Return lblnEsValido
    End Function
    Private Function FblnCambioPropietario() As Boolean
        Dim lblnCambio = False
        If EnuEstadoActualizacion <> EnuEstadoObjetoDef.EnuConsultando Then
            For Each lobjProp As ClsPropietario In ColPropietarios
                If lobjProp.EnuEstadoActualizacion <> EnuEstadoObjetoDef.EnuConsultando Then
                    lblnCambio = lobjProp.BlnTengoCambios
                    If lblnCambio Then Exit For
                End If
            Next
        End If
        Return lblnCambio
    End Function
#End Region
#Region "Procedimientos del objeto"
    ''' <summary>
    ''' Cambia el propietario de los predios agrupados cuando el propietario del predio agrupador
    ''' cambia.
    ''' </summary>
    ''' <summary>
    ''' Cambia el predio agrupador de los predios agrupados cuando el predio agrupador del predio
    ''' agupador cambia.
    ''' </summary>
    Private Sub SActuaPredAgruAPrediosAgrupados()
        Dim lstrIdPredAgru = ObjIdPredioAgrupadorStr.ToString()
        Dim lstrIdPredAgruAnt = ObjIdPredioAgrupadorStr.ObjValorOriginal.ToString()
        Dim lcolCamposCambio As New Collection From {
            ClsIdPredioAgrupadorStr.SstrNombreCampoBd
        }
        Dim lcolDatosNuevos As New Collection From {
            {lstrIdPredAgru, ClsIdPredioAgrupadorStr.SstrNombreCampoBd}
        }
        Dim lcolCamposRef As New Collection From {
            StrCampoCarpeta, StrCampoCentroUtil,
            ClsIdPredioAgrupadorStr.SstrNombreCampoBd
        }
        Dim lcolDatosRef As New Collection From {
            {GshrIdCarpeta, StrCampoCarpeta},
            {GshrIdCentroUtil, StrCampoCentroUtil},
            {lstrIdPredAgruAnt, ClsIdPredioAgrupadorStr.SstrNombreCampoBd}
        }
        GobjPanDat.SActualiceRegistro(SstrNombreTabla, lcolCamposCambio, lcolDatosNuevos,
                lcolCamposRef, lcolDatosRef)

    End Sub
    Private Sub SActualicePrediosAgrupados()
        For Each lobjPredio As ClsPredio In ColPrediosAgrupados
            If lobjPredio.ObjIdPredioStr.ObjValorPro <>
                    lobjPredio.ObjIdPredioAgrupadorStr.ObjValorPro Then
                If ObjIdTipoDestinatarioFacturaByt.ObjValorPro <>
                        lobjPredio.ObjIdTipoDestinatarioFacturaByt.ObjValorPro Then
                    If lobjPredio.EnuEstadoActualizacion =
                            EnuEstadoObjetoDef.EnuConsultando Then
                        lobjPredio.EnuEstadoActualizacion =
                                EnuEstadoObjetoDef.EnuModificando
                    End If
                    lobjPredio.ObjIdTipoDestinatarioFacturaByt.ObjValorPro =
                            ObjIdTipoDestinatarioFacturaByt.ObjValorPro
                End If
                If ObjIdClienteArrendatarioDbl.ObjValorPro <>
                        lobjPredio.ObjIdClienteArrendatarioDbl.ObjValorPro Then
                    If lobjPredio.EnuEstadoActualizacion =
                            EnuEstadoObjetoDef.EnuConsultando Then
                        lobjPredio.EnuEstadoActualizacion =
                                EnuEstadoObjetoDef.EnuModificando
                    End If
                    lobjPredio.ObjIdClienteArrendatarioDbl.ObjValorPro =
                            ObjIdClienteArrendatarioDbl.ObjValorPro
                End If
                If lobjPredio.ObjNoConsolidarItemsFacBln.ObjValorPro <>
                        ObjNoConsolidarItemsFacBln.ObjValorPro Then
                    If lobjPredio.EnuEstadoActualizacion =
                            EnuEstadoObjetoDef.EnuConsultando Then
                        lobjPredio.EnuEstadoActualizacion =
                                EnuEstadoObjetoDef.EnuModificando
                    End If
                    lobjPredio.ObjNoConsolidarItemsFacBln.ObjValorPro =
                        ObjNoConsolidarItemsFacBln.ObjValorPro
                End If
                If lobjPredio.EnuEstadoActualizacion <>
                        EnuEstadoObjetoDef.EnuConsultando Then
                    lobjPredio.SActualice(True)
                End If
            End If
        Next
    End Sub
    Private Sub SActualicePropPredios()
        For Each lobjPredio As ClsPredio In ColPrediosAgrupados
            If lobjPredio.ObjIdPredioStr.ObjValorPro <> ObjIdPredioAgrupadorStr.ObjValorPro Then
                lobjPredio.SActualiceProp(ColPropietarios)
            End If
        Next
    End Sub
    Friend Sub SActualiceProp(acolNewProp As Collection)
        EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
        For Each lobjProp As ClsPropietario In ColPropietarios
            If Not lobjProp.FblnSuprimio Then
                Throw New ErrorInesperadoPanLException("Eliminando propietario")
            End If
        Next
        McolPropietarios = Nothing
        Dim ldrwNewProp As DataRow, ldtbProp = FdtbPropietarios()
        For Each lobjProp As ClsPropietario In acolNewProp
            ldrwNewProp = ldtbProp.NewRow
            Dim lobjNewProp As New ClsPropietario(Me, ldrwNewProp)
            lobjNewProp.SCreeObj(Nothing)
            lobjNewProp.ObjIdCarpeta_PropShr.ObjValorPro = GshrIdCarpeta
            lobjNewProp.ObjIdCentroUtil_PropShr.ObjValorPro = GshrIdCentroUtil
            lobjNewProp.ObjIdPredio_PropStr.ObjValorPro = ObjIdPredioStr.ObjValorPro
            lobjNewProp.ObjIdCliente_PropDbl.ObjValorPro = lobjProp.ObjIdCliente_PropDbl.ObjValorPro
            lobjNewProp.ObjPorcentajePartiDbl.ObjValorPro = lobjProp.ObjPorcentajePartiDbl.ObjValorPro
            ColPropietarios.Add(lobjNewProp, lobjNewProp.ObjIdCliente_PropDbl.ToString())
        Next
        SActualice(True)
    End Sub
    Private Sub SPueblePrediosAgrupados()
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdPredioAgrupadorStr.SstrNombreCampoBd & " = '" &
                ObjIdPredioAgrupadorStr.ObjValorPro & "'"
        Dim ldrwPrediosAgrupados() = ClsPanorama.FdrwDataRow(SstrNombreTabla, {ClsIdPredioStr.SstrNombreCampoBd},
                {{ClsIdPredioStr.SstrNombreCampoBd, "ASC"}}, lstrFiltro)
        If ldrwPrediosAgrupados.Length > 0 Then
            ReDim MstrPrediosAgrupados(ldrwPrediosAgrupados.Length)
            For i = 0 To ldrwPrediosAgrupados.Length - 1
                MstrPrediosAgrupados(i) = ldrwPrediosAgrupados(i)(0)
            Next
        Else
            If Not FblnEstaVacioOrigenDatos() Then
                Throw New ErrorInesperadoPanLException("No hay Predios Agrupados")
            End If
        End If
    End Sub
    Private Sub SPuebleTablaPrediosAgrupados()
        If IsNothing(MdtbPrediosAgrupados) Then
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsIdPredioAgrupadorStr.SstrNombreCampoBd & " = '" &
                    ObjIdPredioAgrupadorStr.ObjValorPro & "' AND '" &
                    ClsIdPredioStr.SstrNombreCampoBd & "' <> '" &
                    ClsIdPredioAgrupadorStr.SstrNombreCampoBd & "'"
            Dim lstrIndice = {{StrCampoCarpeta, "ASC"},
                              {StrCampoCentroUtil, "ASC"},
                              {ClsIdPredioAgrupadorStr.SstrNombreCampoBd, "ASC"},
                              {ClsIdPredioStr.SstrNombreCampoBd, "ASC"}}
            MdtbPrediosAgrupados = ClsPanorama.FdtbDataTable(SstrNombreTabla, {"*"}, lstrIndice, lstrFiltro)
        End If
    End Sub
    Friend Function FdtbFacturasPredio(aenuEstadoFactura As EnuEstadoFacturaDef) As DataTable
        Dim ldtbFacturasPredi As DataTable
        Dim lstrTablaPri = ClsItemFactura.SstrNombreTabla
        Dim lstrTableSec = ClsFactura.SstrNombreTabla
        Dim lstrCamposPri = {ClsPrefijo_ItemFactStr.SstrNombreCampoBd,
                             ClsIdFactura_ItemFactEnt.SstrNombreCampoBd, " ''  AS NroFact",
                             ClsDetalle_ItemFactStr.SstrNombreCampoBd,
                             ClsFechaVencimientoIFDtm.SstrNombreCampoBd,
                             ClsValor_ItemFactDec.SstrNombreCampoBd,
                             ClsIdItemFacturaShr.SstrNombreCampoBd,
                             ("(P." & ClsDebitos_ItemFactDec.SstrNombreCampoBd & " - " & "P." &
                              ClsCreditos_ItemFactDec.SstrNombreCampoBd) & ") AS Saldo"}
        Dim lstrCamposSec = {ClsFechaFacturaDtm.SstrNombreCampoBd}
        Dim lstrCamposRelPri = {StrCampoCarpeta, StrCampoCentroUtil,
                                ClsPrefijo_ItemFactStr.SstrNombreCampoBd,
                                ClsIdFactura_ItemFactEnt.SstrNombreCampoBd}
        Dim lstrCamposRelSec = {StrCampoCarpeta, StrCampoCentroUtil,
                                ClsPrefijo_FactStr.SstrNombreCampoBd, ClsIdFacturaEnt.SstrNombreCampoBd}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion_Pri & " AND " &
                ClsIdPredio_ItemFactStr.SstrNombreCampoBd & " = '" & ObjIdPredioStr.ObjValorPro & "'"
        Dim lstrHoy = ClsPanoramaDat.FstrFechaNormalizada(Today.ToString)
        Dim lstrIndice = {{ClsIdFactura_ItemFactEnt.SstrNombreCampoBd, "ASC"},
                {ClsIdItemFacturaShr.SstrNombreCampoBd, "ASC"}}
        Select Case aenuEstadoFactura
            Case EnuEstadoFacturaDef.EnuNormal
                lstrFiltro &= " AND " & "P." & ClsAnuladoBln.SstrNombreCampoBd & " = False"
            Case EnuEstadoFacturaDef.EnuAnulada
                lstrFiltro &= " AND " & "P." & ClsAnuladoBln.SstrNombreCampoBd & " = True"
            Case EnuEstadoFacturaDef.EnuCancelada
                lstrFiltro &= " AND P." & ClsDebitos_ItemFactDec.SstrNombreCampoBd & " - P." &
                        ClsCreditos_ItemFactDec.SstrNombreCampoBd & " = 0"
            Case EnuEstadoFacturaDef.EnuVencida
                lstrFiltro &= " AND P." & ClsDebitos_ItemFactDec.SstrNombreCampoBd & " - P." &
                        ClsCreditos_ItemFactDec.SstrNombreCampoBd & " > 0" &
                        " AND S." & ClsFechaVencimientoDtm.SstrNombreCampoBd & " < '" & lstrHoy & "'"
        End Select
        ldtbFacturasPredi = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCamposPri,
                lstrTableSec, lstrCamposSec, lstrCamposRelPri, lstrCamposRelSec, lstrIndice,
                lstrFiltro, Array.Empty(Of String), False)
        SComplementeTablaFact(ldtbFacturasPredi)
        Return ldtbFacturasPredi
    End Function
    Private Shared Sub SComplementeTablaFact(adtbFacturasPredio As DataTable)
        Dim lstrPref As String, lentIdFact As Integer
        For Each ldrwFact As DataRow In adtbFacturasPredio.Rows
            lstrPref = ClsPanorama.FobjValorCampo(ldrwFact(0), EnuTipoValor.EnuString)
            lentIdFact = ClsPanorama.FobjValorCampo(ldrwFact(1), EnuTipoValor.EnuInteger)
            Dim lstrNtoFact = ClsPanorama.FstrNumeroDcto(lstrPref, lentIdFact)
            ldrwFact("NroFact") = lstrNtoFact
        Next
    End Sub
    ''' <summary>
    ''' Devuelve el Id del Cliente del Cliente con un Id diferente al pasado en el argumento "adblIdCliente"
    ''' que tiene deuda con el PredioAgrupador actual.
    ''' </summary>
    ''' <param name="adblIdCliente">Id del Cliente referente.</param>
    ''' <returns>Si no existe una Deuda de este Predio con un Cliente diferente al identificado con 
    ''' el Id pasado en el argumento "adblIdCliente", las funcion devuelve cero.</returns>
    ''' <remarks></remarks>
    Friend Function FdblIdClienteDiferenteConDeuda(adblIdCliente As Double) As Double
        Dim ldblIdClienteDif = 0.0
        Dim lstrCamposSelect() = {ClsIdCliente_FactDbl.SstrNombreCampoBd}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsIdCliente_FactDbl.SstrNombreCampoBd &
                " <> " & adblIdCliente & " AND " & ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd & " = '" &
                ObjIdPredioStr.ObjValorPro & "' AND " & ClsDebitos_FactDec.SstrNombreCampoBd & " <> " &
                ClsCreditos_FactDec.SstrNombreCampoBd
        Dim ldtbFact As DataTable = ClsPanorama.FdtbDataTable(ClsFactura.SstrNombreTabla, lstrCamposSelect,
                Nothing, lstrFiltro)
        If ldtbFact.Rows.Count > 0 Then
            ldblIdClienteDif = ClsPanorama.FobjValorCampo(ldtbFact.Rows(0)(0), EnuTipoValor.EnuDouble)
        End If
        Return ldblIdClienteDif
    End Function
    ''' <summary>
    ''' Devuelve un ArrayList con el Id de los Clientes que tienen o han tenido deudas con
    ''' este predio como predio agrupador. 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function FstrClientesDelPredio() As ArrayList
        Dim ldblIdClientes As ArrayList
        GobjPanDat.SControleProcesoObj(True)
        Dim ldtbClientesDelPredio = FdtbClientesDelPredio()
        Dim ldblIdCliente As Double
        ldblIdClientes = New ArrayList
        For Each ldrwCliente As DataRow In ldtbClientesDelPredio.Rows
            ldblIdCliente = ClsPanorama.FobjValorCampo(ldrwCliente(0), EnuTipoValor.EnuDouble)
            ldblIdClientes.Add(ldblIdCliente)
        Next
        For Each lobjProp As ClsPropietario In ColPropietarios
            If Not ldblIdClientes.Contains(lobjProp.ObjIdCliente_PropDbl.ObjValorPro) Then
                ldblIdClientes.Add(lobjProp.ObjIdCliente_PropDbl.ObjValorPro)
            End If
        Next
        GobjPanDat.SControleProcesoObj(False)
        Return ldblIdClientes
    End Function
    Private Function FdtbClientesDelPredio() As DataTable
        Dim lstrCamposSelect = {ClsIdCliente_FactDbl.SstrNombreCampoBd,
                "SUM(" & ClsDebitos_FactDec.SstrNombreCampoBd & " - " &
                ClsCreditos_FactDec.SstrNombreCampoBd & ") AS Saldo"}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd & " = '" & ObjIdPredioStr.ObjValorPro &
                "'"
        Dim ldtbCliePre = ClsPanorama.FdtbDataTable(ClsFactura.SstrNombreTabla, lstrCamposSelect,
                {{"Saldo", "DESC"}}, lstrFiltro, False, {ClsIdCliente_FactDbl.SstrNombreCampoBd})
        Return ldtbCliePre
    End Function
    Friend Function FdecSaldoDeudaPredio(astrIdPredio As String) As Decimal
        Dim lstrTabla = ClsItemFactura.SstrNombreTabla
        Dim lstrCampsoSelect As String() = {"SUM(" & ClsDebitos_ItemFactDec.SstrNombreCampoBd &
                " - " & ClsCreditos_ItemFactDec.SstrNombreCampoBd & ") AS Saldo"}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdPredioStr.SstrNombreCampoBd & " = '" & astrIdPredio & "'"
        Dim ldtbRes = ClsPanorama.FdtbDataTable(lstrTabla, lstrCampsoSelect, {{"", ""}},
                lstrFiltro)
        Dim ldecSaldo As Decimal = ClsPanorama.FobjValorCampo(ldtbRes.Rows(0)(0),
                    EnuTipoValor.EnuDecimal)
        Return ldecSaldo
    End Function
    Friend Function FblnEstaPazYSalvo(ByRef astrMens As String) As Boolean
        Dim lstrPredioNoPazySalvo As String() = {}, i = -1
        Dim lstrIdPredio As String, lblnEstaPazYSalvo = True
        If ObjIdPredioStr.ObjValorPro = ObjIdPredioAgrupadorStr.ObjValorPro Then
            For Each lobjPredio As ClsPredio In ColPrediosAgrupados
                lstrIdPredio = lobjPredio.ObjIdPredioStr.ToString()
                If FdecSaldoDeudaPredio(lstrIdPredio) > 0 Then
                    i += 1
                    ReDim Preserve lstrPredioNoPazySalvo(i)
                    lstrPredioNoPazySalvo(i) = lstrIdPredio
                End If
            Next
            If lstrPredioNoPazySalvo.Length > 0 Then
                Dim lblnUltimo As Boolean, lstrPredios = String.Empty
                For i = 0 To lstrPredioNoPazySalvo.Length - 1
                    lblnUltimo = i = lstrPredioNoPazySalvo.Length - 1
                    lstrPredios &= lstrPredioNoPazySalvo(i)
                    If Not lblnUltimo Then
                        lstrPredios &= ", "
                    End If
                Next
                astrMens = If(i = 1, "El Predio """ & lstrPredios & """ no está a Paz y Salvo!",
                                "Los Predios """ & lstrPredios & """ no están a Paz y Salvo!")
                lblnEstaPazYSalvo = False
            End If
        Else
            lstrIdPredio = ObjIdPredioStr.ObjValorPro
            If FdecSaldoDeudaPredio(lstrIdPredio) = 0 Then
                lblnEstaPazYSalvo = True
            Else
                astrMens = "El presente Predio no está a Paz y Salvo!"
            End If
        End If
        Return lblnEstaPazYSalvo
    End Function
    Friend Function FblnPropsValidos() As Boolean
        Dim lcolPropPred As Collection
        Dim lblnEsValido = ObjIdPredioAgrupadorStr.ObjValorPro =
                ObjIdPredioStr.ObjValorPro
        If Not lblnEsValido Then
            Dim lstrIdCliente As String, lobjPropPredAgr As ClsPropietario
            Dim ldblPorParCli As Double
            Dim lcolPropPredAgr As Collection = ObjPredioAgrupador.ColPropietarios
            lcolPropPred = ColPropietarios
            lblnEsValido = lcolPropPredAgr.Count = lcolPropPred.Count
            If Not lblnEsValido Then Return lblnEsValido
            For Each lobjProp As ClsPropietario In lcolPropPredAgr
                lstrIdCliente = lobjProp.ObjIdCliente_PropDbl.ToString()
                lblnEsValido = lcolPropPredAgr.Contains(lstrIdCliente)
                If Not lblnEsValido Then Exit For
                lobjPropPredAgr = lcolPropPredAgr(lstrIdCliente)
                ldblPorParCli = lobjProp.ObjPorcentajePartiDbl.ObjValorPro
                lblnEsValido = lobjPropPredAgr.ObjPorcentajePartiDbl.ObjValorPro =
                        ldblPorParCli
                If Not lblnEsValido Then Exit For
            Next
        End If
        Return lblnEsValido
    End Function
    Friend Sub SModifiqueParaEstado()
        HblnEsAnulable = False
        HblnEsCreable = False
        HblnEsSuprimible = False
        HblnEsModificable = False
    End Sub
    Friend Function FdecDesctoPP_Prop(adblIdProp As Double, adecValorItemFac As Decimal) As Decimal
        Dim ldecDsctoPP As Decimal
        Dim ldecDctoCli As Decimal = 0.0
        Dim ldblPorPar As Double
        If GobjParametros.ObjAnoActual.ObjAplicaDsctoPPBln.ObjValorPro Then
            If GobjParametros.ObjAnoActual.ObjTipoDsctoPPByt.ObjValorPro =
                        EnuTipoDsctoPP.EnuValorFijo Then
                ldecDsctoPP = ObjSector.ObjDctoProntoPago_SecDbl.ObjValorPro
            Else
                Dim ldblTasaDsctoPP As Double =
                            ObjSector.ObjDctoProntoPago_SecDbl.ObjValorPro
                ldecDsctoPP = ldblTasaDsctoPP * adecValorItemFac
            End If
            If ObjIdTipoDestinatarioFacturaByt.ObjValorPro =
                    EnuDestinatarioFacturaDef.EnuPropietario Then
                Dim lobjProp As ClsPropietario = ColPropietarios(adblIdProp.ToString())
                ldblPorPar = lobjProp.ObjPorcentajePartiDbl.ObjValorPro
                ldecDctoCli = ClsOrionCop.FdecValorRedondeado(ldecDsctoPP *
                                ldblPorPar)
            Else
                ldblPorPar = 1
                ldecDctoCli = ClsOrionCop.FdecValorRedondeado(ldecDsctoPP *
                                ldblPorPar)
            End If
        End If
        Return ldecDctoCli
    End Function
    Friend Function FshrIdSerAdminContribuye() As Short
        Dim lshrIdModContr = FShrIdModuloContribuye(), lshrIdServicio = 0S
        For Each lobjSerAdm As ClsServicio In GobjParametros.ObjAnoActual.ColServiciosAno
            If Not lobjSerAdm.ObjEsAjusteBln.ObjValorPro Then
                If FblnModuloContAServicio(lobjSerAdm, lshrIdModContr) Then
                    lshrIdServicio = lobjSerAdm.ObjIdServicioShr.ObjValorPro
                    Exit For
                End If
            End If
        Next
        Return lshrIdServicio
    End Function
    Private Function FblnModuloContAServicio(aobjServicio As ClsServicio,
            ashrIdModulo As Short) As Boolean
        Dim lblnSi = False
        For Each lobjModSer As ClsModuloServicio In aobjServicio.ColModulosServicio
            If lobjModSer.ObjIdModulo_ModuloServicioShr.ObjValorPro = ashrIdModulo Then
                lblnSi = True
                If lblnSi Then Exit For
            End If
        Next
        Return lblnSi
    End Function
    Private Function FShrIdModuloContribuye() As Short
        Dim lshrIdModulo As Short, lobjModuloContr As ClsModuloContribucion
        Dim lstrTabla = ClsSectorModulo.SstrNombreTabla
        Dim lstrCampSel As String() = {ClsIdModulo_SectorModuloShr.SstrNombreCampoBd}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdSector_SectorModuloShr.SstrNombreCampoBd & " = " &
                ObjIdSector_PredioShr.ObjValorPro
        Dim lstrOrden As String(,) = {{ClsIdModulo_SectorModuloShr.SstrNombreCampoBd, "ASC"}}
        Dim ldtbModulosCont = ClsPanorama.FdtbDataTable(lstrTabla, lstrCampSel, lstrOrden, lstrFiltro)
        For Each ldrwModCon As DataRow In ldtbModulosCont.Rows
            lshrIdModulo = ClsPanorama.FobjValorCampo(ldrwModCon(0), EnuTipoValor.EnuShort)
            lobjModuloContr = GobjParametros.ColModulos(lshrIdModulo.ToString())
            If lobjModuloContr.ObjContribuyeCuotaAdminBln.ObjValorPro Then
                Exit For
            End If
        Next
        Return lshrIdModulo
    End Function
    ''' <summary>
    ''' Indica si el string pasado en el argumento ya existe
    ''' </summary>
    ''' <param name="astrRefPago"></param>
    ''' <returns></returns>
    Friend Function FblnExisteRefPago(astrRefPago As String) As Boolean
        Dim lstrTabla = SstrNombreTabla
        Dim lstrCamSel As String() = {"*"}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsReferenciaPagoStr.SstrNombreCampoBd & " = '" & astrRefPago & "'"
        Dim lstrOrden As String(,) = {{"", ""}}
        Dim lstrGrupo As String() = {}
        Dim ldtbRes As DataTable = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamSel, lstrOrden,
                lstrFiltro, False, lstrGrupo)
        Dim lblnExiste As Boolean = ldtbRes.Rows.Count > 0
        Return lblnExiste
    End Function
    ''' <summary>
    ''' Indica si todos los predios agrupadores tienen referencia de pago o 
    ''' si ninguno tiene referencia de pago.
    ''' </summary>
    ''' <returns></returns>
    Friend Shared Function FblnConRefPago() As Boolean
        Dim lstrTabla = SstrNombreTabla
        Dim lstrCampSel As String() = {"COUNT(" & ClsIdPredioStr.SstrNombreCampoBd & ")"}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsIdPredioStr.SstrNombreCampoBd &
                " = " & ClsIdPredioAgrupadorStr.SstrNombreCampoBd & " AND " &
                ClsReferenciaPagoStr.SstrNombreCampoBd & " <> ''"
        Dim lstrOrden As String(,) = {{"", ""}}
        Dim ldtbRes = ClsPanorama.FdtbDataTable(lstrTabla, lstrCampSel, lstrOrden, lstrFiltro,
                False, {})
        Dim lblnConRefPago = ClsPanorama.FobjValorCampo(ldtbRes.Rows(0)(0), EnuTipoValor.EnuByte) > 0
        Return lblnConRefPago
    End Function
#End Region
#Region "Items Programa Facturacion"
    Friend ReadOnly Property ColItemsProgramaFact(ashrIdAno As Short) As Collection
        Get
            If IsNothing(McolItemsProgramaFact) OrElse ashrIdAno <> MshrIdAno Then
                If Not IsNothing(McolItemsProgramaFact) Then
                    McolItemsProgramaFact.Clear()
                Else
                    McolItemsProgramaFact = New Collection
                End If
                SCargueDtbItemsProgFact()
                Dim ldrwItemsprogramaFact() As DataRow = MdtbItemsProgFact.Select()
                If ldrwItemsprogramaFact.Length > 0 Then
                    Dim lshrIdAno As Short
                    Dim lstrKey As String
                    For Each ldrwItemProgFac As DataRow In ldrwItemsprogramaFact
                        lshrIdAno = ClsPanorama.FobjValorCampo(ldrwItemProgFac(
                                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd), EnuTipoValor.EnuShort)
                        If lshrIdAno = ashrIdAno Then
                            Dim lobjItemprogramaFact As New ClsItemProgramaFact(Me, EnuTipoDeudorDef.EnuPredio, ldrwItemProgFac)
                            lobjItemprogramaFact.SLeaValores(True)
                            lstrKey = lobjItemprogramaFact.ObjIdAno_ItemProgramaFactShr.ToString &
                                    "," & lobjItemprogramaFact.ObjIdServicio_ItemProgramaFactShr.ToString
                            McolItemsProgramaFact.Add(lobjItemprogramaFact, lstrKey)
                        End If
                    Next
                    MshrIdAno = ashrIdAno
                End If
            End If
            Return McolItemsProgramaFact
        End Get
    End Property
    ''' <summary>
    ''' Indica si el Predio tiene algún Servicio Permanente programado para ser facturado.
    ''' </summary>
    ''' <remarks></remarks>
    Friend ReadOnly Property BlnTieneItemsPermanentes As Boolean
        Get
            Dim lblnTiene = False
            If IsNothing(McolItemsProgramaFact) Then
                McolItemsProgramaFact = ColItemsProgramaFact(0)
                lblnTiene = (McolItemsProgramaFact.Count > 0)
            End If
            Return lblnTiene
        End Get
    End Property
    ''' <summary>
    ''' Devuelve el objeto Item del Programa de Facturacion perteneciente al predio e 
    ''' identificado por los argumentos. Si no hay Items devuelve un objeto vacio. 
    ''' Si los argumentos son igual a cero devuelve el primer Item del predio y si no existe 
    ''' el objeto solicitado devuelve Nothing.
    ''' </summary>
    ''' <param name="ashrIdAno">Identifica el Año del Item. Si este argumento es cero 
    ''' el Servicio programado
    ''' es un servicio permanente. Si es sero y el argumento 'ashrIdServicio' es cero
    ''' también devuelve el primer 
    ''' Item del Programa del Predio.
    ''' </param>
    ''' <param name="ashrIdServicio">Identifica el Servicio programado. Si este argumento 
    ''' es cero y el 
    ''' argumento 'ashrIdAno'  también es cero devuelve el primer Item del Programa del Predio.
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend ReadOnly Property ObjItemProgramaFact(ashrIdAno As Short,
            ashrIdServicio As Short) As ClsItemProgramaFact
        Get
            Dim lobjItemProgramaFact As ClsItemProgramaFact = Nothing
            SCargueDtbItemsProgFact()
            Dim ldrwItemsprogFact() As DataRow = MdtbItemsProgFact.Select
            Dim lshrIdAnoBuscado As Short
            If ashrIdAno = 0 AndAlso ashrIdServicio = 0 Then
                lshrIdAnoBuscado = GobjParametros.ObjAnoActual.ObjIdAnoShr.ObjValorPro
            Else
                lshrIdAnoBuscado = ashrIdAno
            End If
            If ldrwItemsprogFact.Length > 0 Then
                For Each ldrwItemProFac As DataRow In ldrwItemsprogFact
                    Dim lshrIdAno As Short = ClsPanorama.FobjValorCampo(
                                ldrwItemProFac(ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd),
                                EnuTipoValor.EnuShort)
                    If ashrIdAno = 0 AndAlso ashrIdServicio = 0 Then
                        If lshrIdAno = lshrIdAnoBuscado Then
                            lobjItemProgramaFact = New ClsItemProgramaFact(Me, EnuTipoDeudorDef.EnuPredio,
                                    ldrwItemProFac)
                            lobjItemProgramaFact.SLeaValores(True)
                            Exit For
                        End If
                    Else
                        Dim lshrIdServicio As Short = ClsPanorama.FobjValorCampo(
                                ldrwItemProFac(ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd),
                                EnuTipoValor.EnuShort)
                        If lshrIdAno = lshrIdAnoBuscado AndAlso lshrIdServicio = ashrIdServicio Then
                            lobjItemProgramaFact = New ClsItemProgramaFact(Me, EnuTipoDeudorDef.EnuPredio,
                                    ldrwItemProFac)
                            lobjItemProgramaFact.SLeaValores(True)
                            Exit For
                        End If
                    End If
                Next
            Else
                Dim ldrwItemProFac = MdtbItemsProgFact.NewRow
                lobjItemProgramaFact = New ClsItemProgramaFact(Me, EnuTipoDeudorDef.EnuPredio, ldrwItemProFac)
            End If
            Return lobjItemProgramaFact
        End Get
    End Property
    ''' <summary>
    ''' Devuelve el item del programa de facturación del predio actual identificado por la llave "astrKeyItem"
    ''' la cual esta conformada por el año y el Id del Servicio correspondiente al item separados por 
    ''' una coma
    ''' </summary>
    ''' <param name="astrKeyItem"></param>
    ''' <returns>Si no existe devuelve Null</returns>
    Friend ReadOnly Property ObjItemProgramaFact(astrKeyItem As String) As ClsItemProgramaFact
        Get
            Dim lobjItemProgramaFact As ClsItemProgramaFact = Nothing
            Dim lstrPartesKey() = Split(astrKeyItem, ",")
            Dim lshrIdAno As Short = CType(lstrPartesKey(0), Short)
            If ColItemsProgramaFact(lshrIdAno).Contains(astrKeyItem) Then
                lobjItemProgramaFact = McolItemsProgramaFact(astrKeyItem)
            End If
            Return lobjItemProgramaFact
        End Get
    End Property
    Friend ReadOnly Property ObjNewItemProgramaFact As ClsItemProgramaFact
        Get
            Dim lobjItemProgramaFact As ClsItemProgramaFact = Nothing
            SCargueDtbItemsProgFact()
            If Not IsNothing(MdtbItemsProgFact) Then
                Dim ldrwNewItemProgramafact As DataRow = MdtbItemsProgFact.NewRow
                lobjItemProgramaFact = New ClsItemProgramaFact(Me, EnuTipoDeudorDef.EnuPredio,
                        ldrwNewItemProgramafact)
                With lobjItemProgramaFact
                    .SCreeObj(Nothing)
                    .ObjIdCarpeta_ItemProgramaFactShr.ObjValorPro = GshrIdCarpeta
                    .ObjIdCentroUtil_ItemProgramaFactShr.ObjValorPro = GshrIdCentroUtil
                    .ObjIdAno_ItemProgramaFactShr.ObjValorPro = 0
                    .ObjIdCliente_ItemProgramaFactDbl.ObjValorPro = 0
                    .ObjIdPredio_ItemProgramaFactStr.ObjValorPro = ObjIdPredioStr.ObjValorPro
                End With
            End If
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
    Friend ReadOnly Property DtbItemsProgFac(ashrIdAno As Short) As DataTable
        Get
            MdtbItemsProgFact = Nothing
            SCargueDtbItemsProgFact()
            Dim lstrFiltro = ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " = " & ashrIdAno.ToString
            Dim lshrIdAnoActual As Short = GobjParametros.ObjAnoActual.ObjIdAnoShr.ObjValorPro
            If ashrIdAno = 0 OrElse ashrIdAno >= lshrIdAnoActual Then
                lstrFiltro &= " OR " & ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd & " = 0"
            End If
            Dim ldrwItemsProgFact() As DataRow = MdtbItemsProgFact.Select(lstrFiltro)
            Dim ldtbItemsPorgramaFact As DataTable = MdtbItemsProgFact.Clone
            For Each ldrwItemProgFac As DataRow In ldrwItemsProgFact
                Dim ldrwNewItem As DataRow = ldtbItemsPorgramaFact.NewRow
                For i = 0 To ldtbItemsPorgramaFact.Columns.Count - 1
                    ldrwNewItem(i) = ldrwItemProgFac(i)
                Next
                ldtbItemsPorgramaFact.Rows.Add(ldrwNewItem)
            Next
            ldtbItemsPorgramaFact.AcceptChanges()
            Return ldtbItemsPorgramaFact
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
    ''' <summary>
    ''' Indica si el item del programa de facturacion identificado por el parametro "astrKeyItemProgramaFact"
    ''' esta pendiente de Facturar
    ''' </summary>
    ''' <param name="astrKeyItemProgramaFact">Identifica el item del programa de facturación</param>
    ''' <param name="ablnIncluirPrediosAgrupados">Indica si debe incluir los predios agrupados por este predio,</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function FblnPendienteporFacturar(astrKeyItemProgramaFact As String,
                ablnIncluirPrediosAgrupados As Boolean) As Boolean
        Dim lblnPorFacturar = False
        Dim lobjItemProgramaFact = ObjItemProgramaFact(astrKeyItemProgramaFact)
        If Not IsNothing(lobjItemProgramaFact) Then
            lblnPorFacturar = lobjItemProgramaFact.FblnPendienteDeFacturar
        End If
        If ablnIncluirPrediosAgrupados Then
            If Not lblnPorFacturar AndAlso ObjIdPredioStr.ObjValorPro =
                    ObjIdPredioAgrupadorStr.ObjValorPro Then
                For Each lobjPredio As ClsPredio In ColPrediosAgrupados
                    lblnPorFacturar = lobjPredio.FblnPendienteporFacturar(astrKeyItemProgramaFact,
                            False)
                    If lblnPorFacturar Then Exit For
                Next
            End If
        End If
        Return lblnPorFacturar
    End Function
    Private Sub SCargueDtbItemsProgFact()
        If IsNothing(MdtbItemsProgFact) Then
            GobjPanDat.SControleProcesoObj(True)
            Dim lstrCamposSelect() = {"*", "'*' AS NombreServicio", "'*' AS NombreOrigen"}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsIdPredio_ItemProgramaFactStr.SstrNombreCampoBd & " = '" &
                    ObjIdPredioStr.ToString & "'"
            Dim lstrIndice = {{ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd, "ASC"},
                             {ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd, "ASC"}}
            MdtbItemsProgFact = ClsPanorama.FdtbDataTable(ClsItemProgramaFact.SstrNombreTabla,
                    lstrCamposSelect, lstrIndice, lstrFiltro)
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
    Private Shared Sub SAsigneOrigenEnDtb(adrwItemOrigramaFac As DataRow)
        Dim lstrNombreCampoOrigen = ClsOrigen_ItemProgramaFacByt.SstrNombreCampoBd
        Dim lbytIdOrigen As Byte = ClsPanorama.FobjValorCampo(adrwItemOrigramaFac(lstrNombreCampoOrigen),
                    EnuTipoValor.EnuByte)
        Dim lstrNombreOrigen = ClsOrionCop.FstrNombreDatoConstanteOri(
                EnuGrupoConstantesOriDef.EnuOrigenItemProgramaFact, lbytIdOrigen)
        adrwItemOrigramaFac("NombreOrigen") = lstrNombreOrigen
    End Sub
#End Region
#Region "ItemsFactura"
    Friend ReadOnly Property DecSaldoPredio As Decimal
        Get
            Dim ldecsaldo = 0D
            If IsNothing(McolItemsFactura) OrElse Not MblnSoloVivasColeccion Then
                McolItemsFactura = ColItemsFactura(True)
            End If
            If McolItemsFactura.Count > 0 Then
                For Each lobjItemFactura As ClsItemFactura In McolItemsFactura
                    ldecsaldo += lobjItemFactura.DecDeuda
                Next
            End If
            Return ldecsaldo
        End Get
    End Property
    Friend ReadOnly Property ColItemsFactura(ablnSoloVivas As Boolean) As Collection
        Get
            If IsNothing(McolItemsFactura) OrElse MblnSoloVivasColeccion <> ablnSoloVivas Then
                MblnSoloVivasColeccion = ablnSoloVivas
                McolItemsFactura = New Collection
                SCargueDtbItemsFactura()
                Dim lstrFiltro = String.Empty
                If ablnSoloVivas Then
                    lstrFiltro &= ClsDebitos_ItemFactDec.SstrNombreCampoBd & " <> " &
                            ClsCreditos_ItemFactDec.SstrNombreCampoBd
                End If
                Dim ldrwItemsFactura() As DataRow = MdtbItemsFactura.Select(lstrFiltro)
                If ldrwItemsFactura.Length > 0 Then
                    For Each ldrwItemFactura As DataRow In ldrwItemsFactura
                        Dim lobjItemFactura As New ClsItemFactura(Me, ldrwItemFactura)
                        lobjItemFactura.SLeaValores(True)
                        Dim lstrKey = lobjItemFactura.ObjPrefijo_ItemFactStr.ToString &
                                lobjItemFactura.ObjIdFactura_ItemFactEnt.ToString &
                                lobjItemFactura.ObjIdItemFacturaShr.ToString
                        McolItemsFactura.Add(lobjItemFactura, lstrKey)
                    Next
                End If
            End If
            Return McolItemsFactura
        End Get
    End Property
    Private Sub SCargueDtbItemsFactura()
        If IsNothing(MdtbItemsFactura) Then
            GobjPanDat.SControleProcesoObj(True)
            Dim lstrCamposSelect() = {"*"}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsIdPredio_ItemFactStr.SstrNombreCampoBd & " = '" &
                    ObjIdPredioStr.ToString & "'"
            Dim lstrIndice = {{ClsPrefijo_ItemFactStr.SstrNombreCampoBd, "ASC"},
                             {ClsIdFactura_ItemFactEnt.SstrNombreCampoBd, "ASC"},
                             {ClsIdItemFacturaShr.SstrNombreCampoBd, "ASC"}}
            MdtbItemsFactura = ClsPanorama.FdtbDataTable(ClsItemFactura.SstrNombreTabla, lstrCamposSelect,
                    lstrIndice, lstrFiltro)
            GobjPanDat.SControleProcesoObj(False)
        End If
    End Sub
#End Region
#Region "Propietarios"
    Friend ReadOnly Property ColPropietarios As Collection
        Get
            If McolPropietarios Is Nothing Then
                McolPropietarios = New Collection
                Dim ldtbProps = FdtbPropietarios(), ldblIdClienteProp As Double
                Dim lobjProp As ClsPropietario
                For Each ldrwProp As DataRow In ldtbProps.Rows
                    ldblIdClienteProp = ClsPanorama.FobjValorCampo(ldrwProp(
                            ClsIdCliente_PropDbl.SstrNombreCampoBd), EnuTipoValor.EnuDouble)
                    lobjProp = New ClsPropietario(Me, ldrwProp)
                    lobjProp.SLeaValores(True)
                    McolPropietarios.Add(lobjProp, lobjProp.ObjIdCliente_PropDbl.ToString())
                Next
            End If
            Return McolPropietarios
        End Get
    End Property
    Friend Function FdtbPropietarios() As DataTable
        Dim lstrTablaPri = ClsPropietario.SstrNombreTabla
        Dim lstrTablaSec = ClsCliente.SstrNombreTabla
        Dim lstrCampSelPri As String() = {"*"}
        Dim lstrCampSelSec As String() = {ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCampRelPri As String() = {StrCampoCarpeta,
                StrCampoCentroUtil, ClsIdCliente_PropDbl.SstrNombreCampoBd}
        Dim lstrCampRelsec As String() = {StrCampoCarpeta,
                StrCampoCentroUtil, ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrFiltro = "P." & ClsIdCarpetaCenUtilShr.SstrNombreCampoBd & " = " & GshrIdCarpeta &
                " AND P." & StrCampoCentroUtil & " = " & GshrIdCentroUtil &
                " AND " & ClsIdPredio_PropStr.SstrNombreCampoBd & " = '" &
                ObjIdPredioStr.ToString() & "'"
        Dim lstrOrden As String(,) = {{ClsPorcentajePartiDbl.SstrNombreCampoBd, "DESC"}}
        Dim ldtbProp = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCampSelPri, lstrTablaSec,
                lstrCampSelSec, lstrCampRelPri, lstrCampRelsec, lstrOrden, lstrFiltro, {}, False)
        Return ldtbProp
    End Function
    Friend Function FobjNewPropietario() As ClsPropietario
        Dim ldtbProp As DataTable = FdtbPropietarios()
        Dim ldrwNewProp = ldtbProp.NewRow
        Dim lobjNewProp As New ClsPropietario(Me, ldrwNewProp)
        lobjNewProp.SCreeObj(Nothing)
        lobjNewProp.ObjIdCarpeta_PropShr.ObjValorPro = GshrIdCarpeta
        lobjNewProp.ObjIdCentroUtil_PropShr.ObjValorPro = GshrIdCentroUtil
        lobjNewProp.ObjIdPredio_PropStr.ObjValorPro = ObjIdPredioStr.ObjValorPro
        lobjNewProp.ObjPorcentajePartiDbl.ObjValorPro = 0
        Return lobjNewProp
    End Function
    Friend Sub SAdicioneNewProp(aobjNewProp As ClsPropietario)
        If EnuEstadoActualizacion <> EnuEstadoObjetoDef.EnuConsultando Then
            aobjNewProp.ObjIdPredio_PropStr.ObjValorPro = ObjIdPredioStr.ObjValorPro
            McolPropietarios.Add(aobjNewProp, aobjNewProp.ObjIdCliente_PropDbl.ToString())
        End If
    End Sub
    Friend Sub SActualiceProp(adblIdCliente As Double, adblPorcientoPart As Double)
        Dim lobjProp As ClsPropietario = ColPropietarios(adblIdCliente.ToString())
        If lobjProp.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
            lobjProp.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
        End If
        lobjProp.ObjPorcentajePartiDbl.ObjValorPro = adblPorcientoPart
    End Sub
    Friend Sub SDesvinculeProp(astrIdProp As String)
        If ColPropietarios.Contains(astrIdProp) Then
            ColPropietarios.Remove(astrIdProp)
        End If
    End Sub
    Friend Function FenuEstadoSugeridoDeuda() As EnuEstadoDeudaDef
        Dim lenuEstadoDeu As EnuEstadoDeudaDef = EnuEstadoDeudaDef.None
        Dim lenuEstSugDeuda As EnuEstadoDeudaDef
        For Each lobjProp As ClsPropietario In ColPropietarios
            lenuEstSugDeuda = lobjProp.ObjCliente.FenuEstadoSugeridoDeuda(
                    ObjIdPredioAgrupadorStr.ToString())
            If lenuEstSugDeuda > lenuEstadoDeu Then
                lenuEstadoDeu = lenuEstSugDeuda
            End If
        Next
        If ObjIdTipoDestinatarioFacturaByt.ObjValorPro =
                EnuDestinatarioFacturaDef.EnuArrendatario AndAlso ObjArrendatario IsNot
                        Nothing Then
            lenuEstSugDeuda =
                    ObjArrendatario.FenuEstadoSugeridoDeuda(ObjIdPredioAgrupadorStr.ToString())
            If lenuEstSugDeuda > lenuEstadoDeu Then
                lenuEstadoDeu = lenuEstSugDeuda
            End If
        End If
        If ObjIdEstadoDeuda_PredioByt.ObjValorPro > lenuEstadoDeu Then
            lenuEstadoDeu = ObjIdEstadoDeuda_PredioByt.ObjValorPro
        End If
        Return lenuEstadoDeu
    End Function
    ' Elimina de las base de datos los propietarios desvinculados
    Private Sub SActualicePropietarios(ByRef ablnCambioPro As Boolean)
        Dim lcolPropDefinitivos As New Collection
        For Each lobjProp As ClsPropietario In ColPropietarios
            lcolPropDefinitivos.Add(lobjProp, lobjProp.ObjIdCliente_PropDbl.ToString())
        Next
        McolPropietarios = Nothing
        For Each lobjProp As ClsPropietario In lcolPropDefinitivos
            If Not ColPropietarios.Contains(lobjProp.ObjIdCliente_PropDbl.ToString()) Then
                ablnCambioPro = True
            End If
        Next
        Dim lstrIdProp As String
        For Each lobjProp As ClsPropietario In ColPropietarios
            lstrIdProp = lobjProp.ObjIdCliente_PropDbl.ObjValorPro.ToString()
            If Not lcolPropDefinitivos.Contains(lstrIdProp) Then
                ablnCambioPro = lobjProp.FblnSuprimio()
            End If
        Next
        McolPropietarios.Clear()
        For Each lobjProp As ClsPropietario In lcolPropDefinitivos
            McolPropietarios.Add(lobjProp, lobjProp.ObjIdCliente_PropDbl.ToString())
        Next
    End Sub
    ' Actualiza los propietarios del predio asignando los propietarios del predio agrupador
    Friend Sub SActualicePropDelPredio(aobjPredioAgr As ClsPredio)
        McolPropietarios = New Collection
        For Each lobjProp As ClsPropietario In aobjPredioAgr.ColPropietarios
            Dim lobjNewProp = FobjNewPropietario()
            lobjNewProp.ObjIdCliente_PropDbl.ObjValorPro = lobjProp.ObjIdCliente_PropDbl.ObjValorPro
            lobjNewProp.ObjNombreCompleto_PropStr.ObjValorPro =
                    lobjProp.ObjNombreCompleto_PropStr.ObjValorPro
            lobjNewProp.ObjPorcentajePartiDbl.ObjValorPro =
                    lobjProp.ObjPorcentajePartiDbl.ObjValorPro
            SAdicioneNewProp(lobjNewProp)
        Next
    End Sub
    Friend Function FblnEsValidoPropietarios(astrIdPredioAgr As String) As Boolean
        Dim lobjPreAgr As New ClsPredio(EnuModoInstanciaObjDef.EnuUnico)
        Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, astrIdPredioAgr}
        Dim lstrKey As String, lblnEsValido As Boolean
        lobjPreAgr.SAbra(lobjValorLlave)
        Dim lcolProPreAgr As Collection = lobjPreAgr.ColPropietarios
        For Each LobjProp As ClsPropietario In ColPropietarios
            lstrKey = LobjProp.ObjIdCliente_PropDbl.ToString()
            lblnEsValido = lcolProPreAgr.Contains(lstrKey)
            If Not lblnEsValido Then Exit For
        Next
        Return lblnEsValido
    End Function
    Friend Function FblnPropietariosCambiaron() As Boolean
        Dim lblnSi As Boolean
        For Each lobjPropietario As ClsPropietario In ColPropietarios
            lblnSi = lobjPropietario.EnuEstadoActualizacion <> EnuEstadoObjetoDef.EnuConsultando
            If lblnSi Then Exit For
        Next
        Return lblnSi
    End Function
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsAliasContStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "AliasCont"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Alias Contable"
        HshrLongitud = 50
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 1,
                ShrLongitud, BlnEsRequerido)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(HobjValorPro) Then
            Return HobjValorPro.ToString
        Else
            Return ""
        End If
    End Function
End Class
Friend Class ClsAreaPredioDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Area"
    Private ReadOnly MobjPadre As ClsPredio = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "BaseParticipacionPredio"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = True
        Dim ldblAreaMinima = 0
        If Not BlnLeyendoOrigen Then
            HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, ldblAreaMinima,
                    Decimal.MaxValue, BlnEsRequerido, HenuTipoValor)
            If Not HblnEsValido Then
                If IsNumeric(HobjValorNew) AndAlso HobjValorNew = 0 Then
                    If MobjPadre.ObjIdSector_PredioShr.BlnEsValido Then
                        Dim lblnFactorContMayorCero = False
                        HblnEsValido = GobjParametros.FblnSectorContribuye(
                            MobjPadre.ObjIdSector_PredioShr.ObjValorPro, lblnFactorContMayorCero)
                        If HblnEsValido Then
                            HblnEsValido = Not lblnFactorContMayorCero
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
Friend Class ClsCoeficientePropiedadDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "CoeficientePropiedad"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "CoeficientePropiedad"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = False
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, 1, True)
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
Friend Class ClsComentarioStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Comentario"
        HshrLongitud = 500
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "Comentarios"
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 2, ShrLongitud, BlnEsRequerido)
    End Sub
    Public Overrides Function ToString() As String
        If Not IsNothing(ObjValorPro) Then
            Return HobjValorPro.ToString
        Else
            Return ""
        End If
    End Function
End Class
Friend Class ClsEmailAdiStr
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsPredio = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "EmailAdi"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Email Adicional"
        HshrLongitud = 100
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 5, ShrLongitud, BlnEsRequerido)
        If HblnEsValido And Not String.IsNullOrEmpty(HobjValorNew) Then
            HobjValorNew = HobjValorNew.ToString.Trim
            HblnEsValido = ClsPanorama.FblnEsValidoEMail(HobjValorNew)
            If Not HblnEsValido Then
                HstrMens = "El Email del Predio no es valido!"
                SNotifiqueDatInv()
            Else
                If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                    If Not IsNothing(HobjValorNew) And HblnEsValido Then
                        HobjValorNew = HobjValorNew.ToString.Trim
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
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsFactorPonderaCPDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FactorPonderacionCP"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Factor ponderación CP"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, 2, HblnEsRequerido,
                HenuTipoValor)
        If HblnEsValido Then
            Dim lobjPadre As ClsPredio = ObjPadre
            If lobjPadre.ObjAreaPredioDec.ObjValorPro > 0 Then
                HblnEsValido = HobjValorNew > 0
                If Not HblnEsValido Then
                    HstrMens = "El Factor de ponderación debe ser mayor a cero!"
                    SNotifiqueDatInv()
                End If
            End If
            HobjValorNew = Math.Round(HobjValorNew, 4)
        Else
            If IsNumeric(HobjValorNew) AndAlso HobjValorNew > 2 Then
                HstrMens = "El Factor de Ponderación máximo admitido es 200 %"
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
Friend Class ClsFacturarPorServicio_PreBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FacturarPorServicio"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Facturar por Agrupador de Servicios"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
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
Friend Class ClsIdClienteAdministradorDbl
    Inherits ClsCBPropiedad
    Private MstrNombreCliente As String = String.Empty
    Private MblnExisteCliente = False
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdClienteAdministrador"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = "IdTerceroAdministrador"
        HblnRegistrarLogCambio = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MstrNombreCliente = String.Empty
        MblnExisteCliente = False
        MyBase.SVaciePropiedad()
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCDBLMINTERC, GCDBLMAXTERC,
                BlnEsRequerido, EnuTipoValor)
        MstrNombreCliente = String.Empty
        MblnExisteCliente = False
        If lblnEsValido Then
            If Not (IsNothing(HobjValorNew) OrElse String.IsNullOrEmpty(HobjValorNew) OrElse
                    HobjValorNew = 0) Then
                Dim lobjLlave() = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
                Dim lobjCliente As ClsCliente = ClsOrionCop.FobjCliente(EnuModoInstanciaObjDef.enuUnico)
                lobjCliente.SAbra(lobjLlave)
                MblnExisteCliente = lobjCliente.BlnExiste
                lblnEsValido = MblnExisteCliente
                If lblnEsValido Then
                    MstrNombreCliente = lobjCliente.ObjNombreCompletoStr.ObjValorPro
                Else
                    MstrNombreCliente = String.Empty
                End If
            End If
        End If
        HblnEsValido = lblnEsValido
    End Sub
    Friend ReadOnly Property BlnExisteCliente As Boolean
        Get
            Return MblnExisteCliente
        End Get
    End Property
    Friend ReadOnly Property StrNombreAdministrador
        Get
            If BlnEsValido Then
                Return MstrNombreCliente
            Else
                Return ""
            End If
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
Friend Class ClsIdClienteArrendatarioDbl
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsPredio = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTerceroArrendatario"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdClienteArrendatario"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsRequerido = (MobjPadre.ObjIdTipoDestinatarioFacturaByt.ObjValorPro =
                        EnuDestinatarioFacturaDef.EnuArrendatario)
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCDBLMINTERC, GCDBLMAXTERC,
                HblnEsRequerido, EnuTipoValor)
        If HblnEsValido AndAlso HobjValorNew > 0 Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                Dim lobjVlrLlave() = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
                Dim lobjCliente As ClsCliente = ClsOrionCop.FobjCliente(EnuModoInstanciaObjDef.enuUnico)
                lobjCliente.SAbra(lobjVlrLlave)
                HblnEsValido = lobjCliente.BlnExiste
                If Not HblnEsValido Then
                    HstrMens = "La Id. del Cliente Ingresada no es valida!"
                    SNotifiqueDatInv()
                End If
            End If
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            MobjPadre.ObjIdTipoDestinatarioFacturaByt.SValide()
        End If
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
Friend Class ClsIdClienteRepLegArrendatariodbl
    Inherits ClsCBPropiedad
    Private MstrNombreCliente As String = String.Empty
    Private MblnExisteCliente = False
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdClienteRepLegArrendatario"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = "IdTerceroRepLegArrendatario"
        HblnRegistrarLogCambio = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MstrNombreCliente = String.Empty
        MblnExisteCliente = False
        MyBase.SVaciePropiedad()
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCDBLMINTERC, GCDBLMAXTERC,
                BlnEsRequerido, EnuTipoValor)
        MstrNombreCliente = String.Empty
        MblnExisteCliente = False
        If lblnEsValido Then
            If Not (IsNothing(HobjValorNew) OrElse String.IsNullOrEmpty(HobjValorNew) OrElse
                    HobjValorNew = 0) Then
                Dim lobjLlave() = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
                Dim lobjCliente As ClsCliente = ClsOrionCop.FobjCliente(EnuModoInstanciaObjDef.enuUnico)
                lobjCliente.SAbra(lobjLlave)
                MblnExisteCliente = lobjCliente.BlnExiste
                lblnEsValido = MblnExisteCliente
                If lblnEsValido Then
                    MstrNombreCliente = lobjCliente.ObjNombreCompletoStr.ObjValorPro
                Else
                    MstrNombreCliente = String.Empty
                End If
            End If
        End If
        HblnEsValido = lblnEsValido
    End Sub
    Friend ReadOnly Property BlnExisteCliente As Boolean
        Get
            Return MblnExisteCliente
        End Get
    End Property
    Friend ReadOnly Property StrNombreRepLegalArrendatario
        Get
            If BlnEsValido Then
                Return MstrNombreCliente
            Else
                Return ""
            End If
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
Friend Class ClsIdEstadoDeuda_PredioByt
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
Friend Class ClsIdFichaCatastralStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdFichaCatastral"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "FichaCatastral"
        HshrLongitud = 21
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 3, ShrLongitud, BlnEsRequerido)
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
Friend Class ClsIdMatriculaInmobiliariaStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdMatriculaInmobiliaria"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "MatriculaInmobiliaria"
        HshrLongitud = 18
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
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
Friend Class ClsIdPredioAgrupadorStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdPredioAgrupador"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdPredioAgrupador"
        HenuTipoValor = EnuTipoValor.enuString
        HshrLongitud = 20
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjValorIng As Object = HobjValorNew
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 2, ShrLongitud,
                BlnEsRequerido)
        If HblnEsValido Then
            Dim lobjPadre As ClsPredio = ObjPadre
            Dim lobjValorLlave() = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
            If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                HstrMens = String.Empty
                If lobjPadre.ObjIdPredioStr.ToString.ToUpper <> HobjValorNew.ToString.ToUpper Then
                    If Not lobjPadre.FblnExisteLlave(lobjValorLlave) Then
                        HstrMens = "La Id. del Predio Agrupador ingresada, '" &
                                lobjValorIng.ToString & "', no existe!"
                        HblnEsValido = False
                    Else
                        Dim lobjPredioAgr As ClsPredio =
                                ClsOrionCop.FobjNuevoPredio(EnuModoInstanciaObjDef.enuUnico)
                        lobjPredioAgr.SAbra(lobjValorLlave)
                        HblnEsValido = lobjPredioAgr.ObjIdPredioAgrupadorStr.ObjValorPro =
                                lobjPredioAgr.ObjIdPredioStr.ObjValorPro
                        If Not HblnEsValido Then
                            HstrMens = "El predio '" & lobjValorIng.ToString &
                                    ", no es un predio agrupador!"
                        Else
                            HobjValorNew = lobjPredioAgr.ObjIdPredioStr.ObjValorPro
                            If HobjValorNew <> HobjValorOriginal Then
                                lobjPadre.SActualicePropDelPredio(lobjPredioAgr)
                            End If
                        End If
                    End If
                Else
                    HobjValorNew = lobjPadre.ObjIdPredioStr.ObjValorPro
                End If
                If Not String.IsNullOrEmpty(HstrMens) AndAlso Not GblnImportando Then
                    SNotifiqueDatInv()
                End If
                If HblnEsValido Then
                    If lobjPadre.ObjIdPredioStr.ObjValorPro <> HobjValorNew Then
                        SValidePropietarios(HobjValorNew)
                    End If
                End If
            Else
                HblnEsValido = lobjPadre.FblnExisteLlave(lobjValorLlave)
            End If
        End If
    End Sub
    Private Sub SValidePropietarios(astrIdPredioAgr As String)
        Dim lobjPadre As ClsPredio = ObjPadre
        If Not GblnImportando Then
            HblnEsValido = lobjPadre.FblnEsValidoPropietarios(astrIdPredioAgr)
            If Not HblnEsValido Then
                If String.IsNullOrEmpty(HstrMens) Then
                    HstrMens = "Los propietarios del predio deben ser los mismos del " &
                                    "predio agrupador"
                End If
                SLevanteEveNot("", 0, EnuSeveridadNot.EnuInformacion)
            End If
        End If
    End Sub
    Private Sub ClsIdPredioStr_evnPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        Dim lobjPadre As ClsPredio = ObjPadre
        If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            If Not e.BlnVaciandoObjeto Then
                lobjPadre.ObjIdPredioStr.SValide()
                lobjPadre.ObjReferenciaPagoStr.SValide()
            End If
        End If
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
Friend Class ClsIdPredioStr
    'Herencia
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdPredio"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdPredio"
        HenuTipoValor = EnuTipoValor.enuString
        HshrLongitud = 20
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 2
    End Sub
    Public Overrides Sub SValide()
        HstrMens = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 2, ShrLongitud,
                    BlnEsRequerido)
        If HblnEsValido Then
            If Not BlnLeyendoOrigen Then
                If ObjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                    HblnEsValido = Not HobjValorNew.ToString.Contains(GCSTRPREFPREFACTURA) AndAlso
                        (HobjValorNew <> "")
                    If Not HblnEsValido Then
                        HstrMens = "La cadena '***' está reservada para el Sistema!"
                    End If
                    If HblnEsValido Then
                        Dim lobjValorLlave() = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
                        If ObjPadre.FblnExisteLlave(lobjValorLlave) Then
                            HstrMens = "La Id. del Predio ingresada, '" &
                                HobjValorNew.ToString & "', ya existe!"
                            HblnEsValido = False
                        End If
                    End If
                ElseIf ObjPadre.EnuEstadoActualizacion =
                        EnuEstadoObjetoDef.enuModificando Then
                    HblnEsValido = (HobjValorOriginal = HobjValorNew)
                    If Not HblnEsValido Then
                        HstrMens = "No es permitido cambiar la identidad a objeto alguno!"
                    End If
                End If
            End If
        Else
            If ObjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando AndAlso
                    Not String.IsNullOrEmpty(HobjValorNew) Then
                HstrMens = "El valor ingresado, '" & HobjValorNew.ToString &
                        "', no es valido!"
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
        If Not IsNothing(ObjValorPro) Then
            Return HobjValorPro.ToString
        Else
            Return ""
        End If
    End Function
End Class
Friend Class ClsIdRegistroMercantilStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdRegistroMercantil"
        HshrLongitud = 20
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "IdRegistroMercantil"
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 2, ShrLongitud, BlnEsRequerido)
    End Sub
    Public Overrides Function ToString() As String
        If Not IsNothing(ObjValorPro) Then
            Return HobjValorPro.ToString
        Else
            Return ""
        End If
    End Function
End Class
Friend Class ClsIdSector_PredioShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdSector"
    Private ReadOnly MobjPadre As ClsPredio = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdSectorPredio"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjValorIngresado = HobjValorNew
        Dim lblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Byte.MaxValue,
                BlnEsRequerido, EnuTipoValor)
        If lblnEsValido Then
            Dim lcolSectores = GobjParametros.ColSectores
            lblnEsValido = lcolSectores.Contains(HobjValorNew.ToString)
        Else
            Dim lobjPadre As ClsPredio = ObjPadre
            If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                If lobjValorIngresado IsNot Nothing AndAlso HobjValorOriginal > 0 Then
                    HstrMens = "La Id. del Sector ingresada, '" &
                        lobjValorIngresado.ToString & "', no existe!"
                    SNotifiqueDatInv()
                End If
            End If
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
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
            If BlnCambio Then
                HstrMens = "ATENCION: Al cambiar un Predio de Sector, " &
                        "se deben recalcular los Coeficientes de Propiedad."
                SLevanteEveNot("", 0, EnuSeveridadNot.EnuAdvertencia)
            End If
        End If
    End Sub
End Class
Friend Class ClsIdTipoDestinatarioFacturaByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTipoDestinatarioFactura"
    Private ReadOnly MobjPadre As ClsPredio = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdTipoDestinatarioFactura"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = "IdTipoDestinatarioFactura"
        HblnRegistrarLogCambio = True
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew,
                EnuDestinatarioFacturaDef.EnuPropietario,
                EnuDestinatarioFacturaDef.EnuArrendatario, HblnEsRequerido)
        If HblnEsValido Then
            If HobjValorNew = EnuDestinatarioFacturaDef.EnuArrendatario Then
                HblnEsValido = (MobjPadre.ObjIdClienteArrendatarioDbl.ObjValorPro > 0)
                If Not HblnEsValido Then
                    HstrMens = "Para facturar al Arrendatario, " &
                            "primero es necesario definir quien es el Arrendatario.!"
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
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
            MobjPadre.ObjIdClienteArrendatarioDbl.SValide()
        End If
    End Sub
    Public Overrides Function ToString() As String
        If Not IsNothing(ObjValorPro) Then
            Return HobjValorPro.ToString
        Else
            Return ""
        End If
    End Function
End Class
Friend Class ClsNoConsolidarItemsFacBln
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "NoConsItemsFac"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "No consolida Items Factura"
        HenuTipoValor = EnuTipoValor.enuBoolean
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = False
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
Friend Class ClsNombreComercialStr
    Inherits ClsCBPropiedad
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "NombreComercial"
        HshrLongitud = 40
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = "NombreComercial"
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 2, ShrLongitud, BlnEsRequerido)
    End Sub
    Public Overrides Function ToString() As String
        If Not IsNothing(ObjValorPro) Then
            Return HobjValorPro.ToString
        Else
            Return ""
        End If
    End Function
End Class
Friend Class ClsNombrePredioStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "NombrePredio"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Nombre  Predio"
        HshrLongitud = 35
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 2, ShrLongitud,
                    BlnEsRequerido)
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
Friend Class ClsReferenciaPagoStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "ReferenciaPago"
    Private ReadOnly MobjPadre As ClsPredio = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Referencia  Pago"
        HshrLongitud = 8
        HenuTipoValor = EnuTipoValor.EnuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud,
                    BlnEsRequerido)
        HstrMens = String.Empty
        If HblnEsValido Then
            If ObjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.EnuConsultando Then
                HobjValorNew = HobjValorNew.ToString().Trim
                If MobjPadre.ObjIdPredioStr.ObjValorPro = MobjPadre.ObjIdPredioAgrupadorStr.ObjValorPro Then
                    If Not String.IsNullOrEmpty(HobjValorNew) Then
                        If HobjValorNew <> HobjValorOriginal Then
                            HblnEsValido = Not MobjPadre.FblnExisteRefPago(HobjValorNew)
                            If Not HblnEsValido Then
                                HstrMens = "La Referencia de Pago ingresada ya está asignada a otro Predio!"
                            End If
                        End If
                    End If
                End If
            End If
        Else
            If Not String.IsNullOrEmpty(HobjValorNew.ToString().Trim) Then
                HstrMens = "La referencia de pago debe tener máximo 8 caractéres"
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
        If Not IsNothing(ObjValorPro) Then
            Return HobjValorPro.ToString
        Else
            Return ""
        End If
    End Function
End Class
Friend Class ClsValorServicioIdDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "ValorServicioId"
    Private ReadOnly MobjPadre As ClsPredio = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Valor Servicio Id"
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HenuTipoValor = EnuTipoValor.enuDecimal
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HstrMens = String.Empty
        HblnEsRequerido = GobjParametros.ObjServicioIdActivoBln.ObjValorPro AndAlso
                    (MobjPadre.ObjIdPredioStr.ObjValorPro = MobjPadre.ObjIdPredioAgrupadorStr.ObjValorPro)
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Decimal.MaxValue,
                HblnEsRequerido, EnuTipoValor)
        If HblnEsValido Then
            If Not GobjParametros.ObjServicioIdActivoBln.ObjValorPro AndAlso HobjValorNew > 0 Then
                HblnEsValido = False
                HstrMens = "Cuando no esta activado el Servicio de Identificación el valor debe ser cero!"
                SNotifiqueDatInv()
            ElseIf GobjParametros.ObjServicioIdActivoBln.ObjValorPro AndAlso HobjValorNew > 0 Then
                If MobjPadre.ObjIdPredioStr.ObjValorPro <> MobjPadre.ObjIdPredioAgrupadorStr.ObjValorPro Then
                    HblnEsValido = False
                    HstrMens = "Solo los Predios Agrupadores deben tener Valor del Servicio de Identificación!"
                End If
            End If
        Else
            If (MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando AndAlso
                        HobjValorNew <> 0) Then
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
        If IsNothing(ObjValorPro) Then
            Return ""
        Else
            Return Format(HobjValorPro, "c")
        End If
    End Function
End Class
#End Region
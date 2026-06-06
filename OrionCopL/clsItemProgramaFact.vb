Friend Class ClsItemProgramaFact
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriItemsProgramaFact"
    ' Variables de modulo
    Private MobjServicio_ItemProgramaFact As clsServicio = Nothing
    Private MobjClienteItem As clsCliente = Nothing
    Private MobjPredioItem As clsPredio = Nothing
    Private MobjAno As clsAno = Nothing
    Private ReadOnly MenuTipoDeudor As EnuTipoDeudorDef = EnuTipoDeudorDef.None
#End Region

#Region "Constructores"
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="adrwObjeto">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Friend Sub New(aobjPadre As ClsCBObjetoPan, aenuTipoDeudor As EnuTipoDeudorDef,
                   adrwObjeto As DataRow)
        HobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsAnulable = False
        '
        DrwRegistroActual = adrwObjeto
        DtbTablaColeccion = DrwRegistroActual.Table
        MenuTipoDeudor = aenuTipoDeudor
    End Sub

    ''' <summary>
    ''' Instancia el Objeto en modo unico solo para leer los permisos
    ''' </summary>
    Public Sub New()
        Dim lstrCamposSelect As String()
        HobjPadre = Nothing
        HblnEsCreable = False
        HblnEsModificable = False
        HblnEsSuprimible = False
        HblnEsAnulable = False
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuUnico
        lstrCamposSelect = {"*"}
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
            Return EnuIdClasesPanDef.enuItemProgFact
        End Get
    End Property

    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Item Programa Facturación"
        End Get
    End Property
#End Region

#Region "Propiedades Prop"
    Friend ReadOnly Property ObjCantidadPeriodosShr As New ClsCantidadPeriodosShr(Me)
    Friend ReadOnly Property ObjIdAno_ItemProgramaFactShr As New ClsIdAno_ItemProgramaFactShr(Me)
    Friend ReadOnly Property ObjIdCarpeta_ItemProgramaFactShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_ItemProgramaFactShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdCliente_ItemProgramaFactDbl As New ClsIdCliente_ItemProgramaFactDbl(Me)
    Friend ReadOnly Property ObjIdPredio_ItemProgramaFactStr As New ClsIdPredio_ItemProgramaFactStr(Me)
    Friend ReadOnly Property ObjIdServicio_ItemProgramaFactShr As New ClsIdServicio_ItemProgramaFactShr(Me)
    Friend ReadOnly Property ObjLecturaActual_ItemProgramaFactDec As New ClsLecturaActual_ItemProgramaFactDec(Me)
    Friend ReadOnly Property ObjLecturaAnterior_ItemProgramaFactDec As New ClsLecturaAnterior_ItemProgramaFactDec(Me)
    Friend ReadOnly Property ObjOrdinal_ItemProgramaFact As New ClsOrdinal_ItemProgramaFact(Me)
    Friend ReadOnly Property ObjOrigen_ItemProgramaFacByt As New ClsOrigen_ItemProgramaFacByt(Me)
    Friend ReadOnly Property ObjPeriodoIni_ItemProgStr As New ClsPeriodoIni_ItemProgStr(Me)
    Friend ReadOnly Property ObjSaldo_ItemProgramaFactDec As New ClsSaldo_ItemProgramaFactDec(Me)
    Friend ReadOnly Property ObjValorPeriodo_ItemProgramaFactDec As New ClsValorPeriodo_ItemProgramaFactDec(Me)
    Friend ReadOnly Property ObjValorUnitario_ItemProgramaFactDec As New ClsValorUnitario_ItemProgramaFactDec(Me)

    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjCantidadPeriodosShr)
                HcolPropiedades.Add(ObjIdAno_ItemProgramaFactShr)
                HcolPropiedades.Add(ObjIdCarpeta_ItemProgramaFactShr)
                HcolPropiedades.Add(ObjIdCentroUtil_ItemProgramaFactShr)
                HcolPropiedades.Add(ObjIdCliente_ItemProgramaFactDbl)
                HcolPropiedades.Add(ObjIdPredio_ItemProgramaFactStr)
                HcolPropiedades.Add(ObjIdServicio_ItemProgramaFactShr)
                HcolPropiedades.Add(ObjLecturaActual_ItemProgramaFactDec)
                HcolPropiedades.Add(ObjLecturaAnterior_ItemProgramaFactDec)
                HcolPropiedades.Add(ObjOrdinal_ItemProgramaFact)
                HcolPropiedades.Add(ObjOrigen_ItemProgramaFacByt)
                HcolPropiedades.Add(ObjPeriodoIni_ItemProgStr)
                HcolPropiedades.Add(ObjSaldo_ItemProgramaFactDec)
                HcolPropiedades.Add(ObjValorPeriodo_ItemProgramaFactDec)
                HcolPropiedades.Add(ObjValorUnitario_ItemProgramaFactDec)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region

#Region "Otras propiedades"
    Friend Property BlnActualizando As Boolean = False
    ''' <summary>
    ''' Devuelve la fecha de la última facturar a generar.
    ''' </summary>
    Friend ReadOnly Property DtmFechaUltimaFactura As Date
        Get
            Dim lentAno = CType(ObjPeriodoIni_ItemProgStr.ObjValorPro.ToString.Substring(0, 4), Integer)
            Dim lentMes = CType(Right(ObjPeriodoIni_ItemProgStr.ObjValorPro.ToString, 2), Integer)
            Dim lentDia As Integer = ObjServicio_ItemProgramaFact.ObjDiaFacturaShr.ObjValorPro
            Dim ldtmDechaFin As Date = DateSerial(lentAno, lentMes, lentDia)
            Return ldtmDechaFin
        End Get
    End Property

    Friend ReadOnly Property DecSaldo As Decimal
        Get
            Return ObjSaldo_ItemProgramaFactDec.ObjValorPro
        End Get
    End Property

    Friend ReadOnly Property DecValorTotal As Decimal
        Get
            Return ObjValorPeriodo_ItemProgramaFactDec.ObjValorPro * ObjCantidadPeriodosShr.ObjValorPro
        End Get
    End Property

    Friend ReadOnly Property ObjServicio_ItemProgramaFact As ClsServicio
        Get
            If IsNothing(MobjServicio_ItemProgramaFact) Then
                If ObjIdServicio_ItemProgramaFactShr.BlnEsValido Then
                    Dim lobjValorLlave As Object()
                    Dim lshrIdAno As Short
                    If ObjAno_ItemProgramaFact Is Nothing Then
                        lshrIdAno = 0
                    Else
                        lshrIdAno = ObjAno_ItemProgramaFact.ObjIdAnoShr.ObjValorPro
                    End If
                    MobjServicio_ItemProgramaFact = New ClsServicio(ObjAno_ItemProgramaFact,
                                EnuModoInstanciaObjDef.enuUnico)
                    lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lshrIdAno,
                                ObjIdServicio_ItemProgramaFactShr.ToString}
                    MobjServicio_ItemProgramaFact.SAbra(lobjValorLlave)
                End If
            End If
            Return MobjServicio_ItemProgramaFact
        End Get
    End Property

    Friend ReadOnly Property DblTarifaIva_Servicio As Double
        Get
            MobjServicio_ItemProgramaFact = Nothing
            MobjServicio_ItemProgramaFact = ObjServicio_ItemProgramaFact
            Dim ldblTarIva = 0.0
            If Not IsNothing(MobjServicio_ItemProgramaFact) Then
                ldblTarIva = ObjServicio_ItemProgramaFact.ObjTarifaIvaDbl.ObjValorPro
            End If
            Return ldblTarIva
        End Get
    End Property

    Friend ReadOnly Property ObjAno_ItemProgramaFact As ClsAno
        Get
            If IsNothing(MobjAno) Then
                If ObjIdAno_ItemProgramaFactShr.BlnEsValido AndAlso
                        ObjIdAno_ItemProgramaFactShr.ObjValorPro > 0 Then
                    MobjAno = GobjParametros.ColAnos(ObjIdAno_ItemProgramaFactShr.ToString)
                End If
            End If
            Return MobjAno
        End Get
    End Property

    Friend ReadOnly Property EnuTipoDeudor As EnuTipoDeudorDef
        Get
            Return MenuTipoDeudor
        End Get
    End Property

    Friend ReadOnly Property ObjClienteItem As ClsCliente
        Get
            If IsNothing(MobjClienteItem) Then
                If MenuTipoDeudor = EnuTipoDeudorDef.EnuCliente Then
                    MobjClienteItem = New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
                    MobjClienteItem.SAbra({GshrIdCarpeta, GshrIdCentroUtil, ObjIdCliente_ItemProgramaFactDbl.ObjValorPro})
                End If
            End If
            Return MobjClienteItem
        End Get
    End Property

    Friend ReadOnly Property ObjPredioItem As ClsPredio
        Get
            If IsNothing(MobjPredioItem) Then
                If MenuTipoDeudor = EnuTipoDeudorDef.EnuPredio Then
                    MobjPredioItem = New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
                    MobjPredioItem.SAbra({GshrIdCarpeta, GshrIdCentroUtil, ObjIdPredio_ItemProgramaFactStr.ObjValorPro})
                End If
            End If
            Return MobjPredioItem
        End Get
    End Property

    Friend ReadOnly Property BlnEsServicioConsumo As Boolean
        Get
            Dim lblnEsSerCons = ObjLecturaAnterior_ItemProgramaFactDec.ObjValorPro <> 0 AndAlso
                    ObjLecturaActual_ItemProgramaFactDec.ObjValorPro <> 0 AndAlso
                    ObjValorUnitario_ItemProgramaFactDec.ObjValorPro <> 0
            Return lblnEsSerCons
        End Get
    End Property
#End Region
#End Region

#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SVacie()
        MobjAno = Nothing
        MobjClienteItem = Nothing
        MobjPredioItem = Nothing
        MobjServicio_ItemProgramaFact = Nothing
        MyBase.SVacie()
    End Sub

    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        Dim lshrIdAno As Short, lshrIdServicio As Short, lstrKey As String
        Try
            GobjPanDat.SControleProcesoObj(True)
            lshrIdAno = ObjIdAno_ItemProgramaFactShr.ObjValorPro
            lshrIdServicio = ObjIdServicio_ItemProgramaFactShr.ObjValorPro
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
                If GblnImportando Then
                    lstrKey = lshrIdAno.ToString & "," & lshrIdServicio.ToString & "," &
                            ObjCantidadPeriodosShr.ToString & "," &
                            ObjPeriodoIni_ItemProgStr.ToString
                    If Not GarlServiciosImportados.Contains(lstrKey) Then
                        GarlServiciosImportados.Add(lstrKey)
                    End If
                    If lshrIdAno > 0 Then
                        SAdecueValores()
                    End If
                End If
                ObjIdCarpeta_ItemProgramaFactShr.ObjValorPro = GshrIdCarpeta
                ObjIdCentroUtil_ItemProgramaFactShr.ObjValorPro = GshrIdCentroUtil
                ObjIdAno_ItemProgramaFactShr.ObjValorPro = lshrIdAno
                ObjIdServicio_ItemProgramaFactShr.ObjValorPro = lshrIdServicio
                SActualiceSaldoActual()
                If IsNothing(ObjOrdinal_ItemProgramaFact.ObjValorPro) OrElse
                        ObjOrdinal_ItemProgramaFact.ObjValorPro = 0 Then
                    SNumereObj()
                End If
            End If
            DtbTablaColeccion.Rows.Clear()
            MyBase.SActualice(ablnExigeRequeridos)
            If GblnImportando AndAlso BlnImportoUltimo Then
                SActualiceServiciosImportado(lshrIdAno)
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
            GobjPanDat.SControleProcesoObj(False)
        End Try
    End Sub

    Friend Overrides Function FblnEsSuprimible() As Boolean
        Dim lblnEsSuprimible As Boolean =
                Not ObjServicio_ItemProgramaFact.ObjGeneraProgramBln.ObjValorPro
        Return lblnEsSuprimible
    End Function

    Protected Overrides Sub SPrepareParaImportacion()
        MyBase.SPrepareParaImportacion()
        DrwRegistroActual = DrwRegistroActual.Table.NewRow
        ObjOrigen_ItemProgramaFacByt.ObjValorPro =
                EnuOrigenItemProgramaFactDef.EnuImportado
        ObjSaldo_ItemProgramaFactDec.ObjValorPro = 0
        If GarlServiciosImportados Is Nothing Then
            GarlServiciosImportados = New ArrayList
        End If
    End Sub

    Friend Overrides Function FblnEsCreable(aobjValorLlave() As Object) As Boolean
        Dim lblnEsCreable = False
        If GblnImportando Then
            If IsNothing(aobjValorLlave) Then
                lblnEsCreable = True
            Else
                Dim lshrIdAno As Short = ObjIdAno_ItemProgramaFactShr.ObjValorPro
                Dim lshrIdServicio As Short = ObjIdServicio_ItemProgramaFactShr.ObjValorPro
                Dim lstrKey As String = lshrIdAno.ToString & "," & lshrIdServicio.ToString
                If lshrIdAno = 0 Then
                    lblnEsCreable = GobjParametros.ColServiciosPer.Contains(lstrKey)
                Else
                    Dim lobjAno As ClsAno = GobjParametros.ColAnos(lshrIdAno.ToString)
                    lblnEsCreable = lobjAno.ColServiciosAno.Contains(lstrKey)
                End If
            End If
        Else
            If BlnEsCreable Then
                If FblnValorLlaveEsNull(aobjValorLlave) Then
                    lblnEsCreable = True
                Else
                    lblnEsCreable = Not FblnExisteLlave(aobjValorLlave)
                End If
            End If
        End If
        Return lblnEsCreable
    End Function

    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjOrdinal_ItemProgramaFact.ToString
        End Get
    End Property

    Friend Overrides Function FblnSonValidosDatosOrigen(adtbOrigen As DataTable,
            astrColumnasRelacionadas As String(), ablnReinicie As Boolean,
            ByRef astrMens As String) As Boolean
        Dim lblnEsValido As Boolean, i = 0, lstrColumnaOrigen As String
        Dim lbytIdCar As Byte, lbytIdCenutil As Byte, ldblIdTerCliente As Double
        Dim lstrIdPredio As String, ldecValor As Decimal
        Dim lstrPeriDeInicio As String, lentCantPeri As Integer
        Dim lstrPerIniTodos = String.Empty
        Dim lobjTer As New ClsTercero(EnuModoInstanciaObjDef.EnuUnico)
        Dim lobjPred As New ClsPredio(EnuModoInstanciaObjDef.EnuUnico)
        Dim lobjServicio As ClsServicio
        Dim larlServicios As New ArrayList
        Dim lstrKeySer As String, larlKeysSer As New ArrayList, larlDestinatario As New ArrayList
        Static lshrIdAno As Short = -1, lshrIdSer As Short = -1
        Static lblnEsServicioAjuste = False
        Dim lobjValorLlave As Object()
        Dim lblnEsConsumo = FblnEsConsumo(adtbOrigen)
        If ablnReinicie Then
            lshrIdAno = -1
            lshrIdSer = -1
            lblnEsServicioAjuste = False
        End If
        lblnEsValido = adtbOrigen.Rows.Count > 0
        If lblnEsValido Then
            For Each ldrwOrigen As DataRow In adtbOrigen.Rows
                i += 1
                lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                        ClsIdCarpetaShr.SstrNombreCampoBd)
                lbytIdCar = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrColumnaOrigen),
                        EnuTipoValor.EnuString)
                lblnEsValido = lbytIdCar = GshrIdCarpeta
                If Not lblnEsValido Then
                    astrMens = "La Carpeta del Registro " & i.ToString &
                            " no es la Carpeta actual!"
                    Exit For
                End If
                If Not lblnEsConsumo Then
                    lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                            ClsIdCentroUtilShr.SstrNombreCampoBd)
                    lbytIdCenutil = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrColumnaOrigen),
                        EnuTipoValor.EnuString)
                    lblnEsValido = lbytIdCenutil = GshrIdCentroUtil
                    If Not lblnEsValido Then
                        astrMens = "La Copropiedad del Registro " & i.ToString &
                            " no es el actual!"
                        Exit For
                    End If
                End If
                lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                        ClsPeriodoIni_ItemProgStr.SstrNombreCampoBd)
                lstrPeriDeInicio = ClsPanorama.FobjValorCampo(
                        ldrwOrigen(lstrColumnaOrigen), EnuTipoValor.EnuString)
                lblnEsValido = lstrPeriDeInicio.Length = 6 AndAlso IsNumeric(lstrPeriDeInicio)
                If Not lblnEsValido Then
                    astrMens = "El período de inicio en el registro " & i.ToString &
                            " debe numérico y tener una longitud de seis dígitos!"
                    Exit For
                End If
                If String.IsNullOrEmpty(lstrPerIniTodos) Then
                    lstrPerIniTodos = lstrPeriDeInicio
                Else
                    lblnEsValido = lstrPeriDeInicio = lstrPerIniTodos
                    If Not lblnEsValido Then
                        astrMens = "El periodo de inicio en el registro " & i.ToString &
                                " es diferente al periodo de inicio de los registros " &
                                "anteriores!"
                        Exit For
                    End If
                End If
                If lshrIdAno = -1 Then
                    If lblnEsConsumo Then
                        lshrIdAno = 0
                    Else
                        lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd)
                        lshrIdAno = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrColumnaOrigen),
                                EnuTipoValor.EnuShort)
                        lblnEsValido = lshrIdAno >= GobjParametros.ObjAnoActual.ObjIdAnoShr.
                                ObjValorPro OrElse lshrIdAno = 0
                        If Not lblnEsValido Then
                            astrMens = "El año del servicio a importar es anterior al año actual!"
                            Exit For
                        End If
                    End If
                Else
                    If Not lblnEsConsumo Then
                        lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                                ClsIdAno_ItemProgramaFactShr.SstrNombreCampoBd)
                        lblnEsValido = lshrIdAno = ClsPanorama.FobjValorCampo(
                                ldrwOrigen(lstrColumnaOrigen), EnuTipoValor.EnuShort)
                        If Not lblnEsValido Then
                            lshrIdAno = -1
                            astrMens = "El Año debe ser igual para todos los Registros!"
                            Exit For
                        End If
                    End If
                End If
                lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                                ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd)
                lshrIdSer = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrColumnaOrigen),
                    EnuTipoValor.EnuShort)
                lstrKeySer = lshrIdAno.ToString() & "," & lshrIdSer.ToString
                If Not larlKeysSer.Contains(lstrKeySer) Then
                    larlKeysSer.Add(lstrKeySer)
                End If
                If lshrIdAno > -1 Then
                    lstrKeySer = lshrIdAno.ToString & "," & lshrIdSer.ToString
                    lobjServicio = GobjParametros.FobjServicio(lstrKeySer)
                    If lobjServicio Is Nothing Then
                        lblnEsValido = False
                        astrMens = "El Servicio en el Registro " & i.ToString &
                                " no existe!"
                        Exit For
                    Else
                        lblnEsServicioAjuste = lobjServicio.ObjEsAjusteBln.ObjValorPro
                    End If
                    lblnEsValido = lobjServicio.ObjEsFactProgramableBln.ObjValorPro
                    If Not lblnEsValido Then
                        astrMens = "El servicio en el Registro " & i.ToString &
                                " debe ser programable para poder importar " &
                                " los valores a cobrar!"
                        Exit For
                    End If
                    Dim lshrCantidadModCon As Short = lobjServicio.ColModulosServicio.Count
                    If lobjServicio.ColModulosServicio.Count = 0 Then
                        lblnEsValido = lshrIdAno = 0
                        If Not lblnEsValido Then
                            astrMens = "No se han definido los Módulos de Contribución para " &
                                        "el Servicio del registro " & i.ToString
                            Exit For
                        End If
                    End If
                    If lobjServicio.ObjGeneraProgramBln.ObjValorPro AndAlso
                            lshrIdAno > 0 Then
                        lblnEsValido = lobjServicio.ColModulosServicio.Count > 0
                        If Not lblnEsValido AndAlso Not lblnEsServicioAjuste Then
                            astrMens = " Los módulos que contribuyen al servicio en " &
                                "el Registro " & i.ToString & " no han sido definidos!"
                            lblnEsValido = False
                            Exit For
                        Else
                            lblnEsValido = True
                        End If
                    End If
                End If
                lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                        ClsIdPredio_ItemProgramaFactStr.SstrNombreCampoBd)
                lstrIdPredio = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrColumnaOrigen),
                        EnuTipoValor.EnuString)
                lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                        ClsIdCliente_ItemProgramaFactDbl.SstrNombreCampoBd)
                ldblIdTerCliente = ClsPanorama.FobjValorCampo(
                        ldrwOrigen(lstrColumnaOrigen), EnuTipoValor.EnuDouble)
                If lblnEsConsumo Then
                    lentCantPeri = 1
                Else
                    lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                            ClsCantidadPeriodosShr.SstrNombreCampoBd)
                    lentCantPeri = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrColumnaOrigen),
                        EnuTipoValor.EnuInteger)
                End If
                lblnEsValido = FblnEsValidoSer(lshrIdAno, lshrIdSer, i, astrMens)
                If Not lblnEsValido Then Exit For
                If lshrIdAno > 0 Then
                    lblnEsValido = lentCantPeri <= 12
                    If Not lblnEsValido Then
                        astrMens = "La cantidad de Períodos en el Registro " &
                                i.ToString() & " no es válida!"
                        Exit For
                    End If
                End If
                lblnEsValido = ClsPanorama.FblnEsValidoNumero(ldblIdTerCliente, 0,
                        Double.MaxValue, True, EnuTipoValor.EnuDouble)
                If Not lblnEsValido Then
                    astrMens = "La Identificación del Cliente en el Registro " &
                        i.ToString & " no es válida!"
                    Exit For
                End If
                If ldblIdTerCliente > 0 Then
                    lobjValorLlave = {ldblIdTerCliente}
                    lobjTer.SAbra(lobjValorLlave)
                    lblnEsValido = lobjTer.BlnExiste
                    If Not lblnEsValido Then
                        astrMens = "El Tercero correspondiente al registro " &
                        i.ToString & " no existe!"
                        Exit For
                    End If
                End If
                If Not larlDestinatario.Contains(lstrIdPredio & ldblIdTerCliente.ToString) Then
                    larlDestinatario.Add(lstrIdPredio & ldblIdTerCliente.ToString)
                Else
                    lblnEsValido = False
                    astrMens = "El destinatario del registro " & i.ToString() & " está repetido!"
                    Exit For
                End If
                If Not String.IsNullOrEmpty(lstrIdPredio) Then
                    lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrIdPredio}
                    lobjPred.SAbra(lobjValorLlave)
                    lblnEsValido = lobjPred.BlnExiste
                    If Not lblnEsValido Then
                        astrMens = "El Predio del registro " & i.ToString & " no existe!"
                        Exit For
                    End If
                End If
                lblnEsValido = FblnEsDestiVali(ldblIdTerCliente, lstrIdPredio, i, astrMens)
                If Not lblnEsValido Then Exit For
                If lshrIdAno > 0 AndAlso Not lblnEsServicioAjuste Then
                    lblnEsValido = lstrPeriDeInicio = lshrIdAno.ToString & "01"
                Else
                    Dim lstrIdPerioAct = GobjParametros.ObjAnoActual.StrIdPeriodoActual
                    lblnEsValido = lstrPeriDeInicio >= lstrIdPerioAct
                End If
                If Not lblnEsValido Then
                    astrMens = "El Período de inicio del registro " & i.ToString &
                        " no es válido!"
                    Exit For
                End If
                If lblnEsConsumo Then
                    lblnEsValido = FblnEsValidoConsumo(ldrwOrigen, ldecValor, astrMens)
                    If Not lblnEsValido Then
                        astrMens = "En el registro " & i.ToString & astrMens
                        Exit For
                    End If
                Else
                    lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                            ClsValorPeriodo_ItemProgramaFactDec.SstrNombreCampoBd)
                    ldecValor = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrColumnaOrigen),
                            EnuTipoValor.EnuDecimal)
                    lblnEsValido = ldecValor - Int(ldecValor) = 0
                    If Not lblnEsValido Then
                        astrMens = "El valor del periodo en el registro " & i.ToString &
                            ", debe ser sin centavos!"
                        Exit For
                    Else
                        lblnEsValido = ldecValor >= 0
                        If Not lblnEsValido Then
                            astrMens = "El valor del periodo en el registro " & i.ToString &
                            " es negativo!"
                            Exit For
                        End If
                    End If
                End If
            Next
            If larlKeysSer.Count > 0 AndAlso ablnReinicie And lblnEsValido Then
                For Each lstrKey As String In larlKeysSer
                    ClsCalculosServicios.SElimineItemsServicio(lstrKey)
                Next
            End If
        Else
            astrMens = "La tabla seleccionada no tiene registros!"
        End If
        Return lblnEsValido
    End Function
#End Region

#Region "Procedimientos del objeto"
    ''' <summary>
    ''' Indica si el item del programa de facturación esta pendiente de facturarse.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function FblnPendienteDeFacturar() As Boolean
        Dim lblnPorFact = False
        If ObjValorPeriodo_ItemProgramaFactDec.ObjValorPro > 0 Then
            Dim lshrPeriodosPorFact As Short = DecSaldo / ObjValorPeriodo_ItemProgramaFactDec.ObjValorPro
            If lshrPeriodosPorFact > 0 Then
                Dim lstrPeriodoActual = GobjParametros.ObjAnoActual.StrIdPeriodoActual
                Dim lshrPeriodosFact As Short = ObjCantidadPeriodosShr.ObjValorPro - lshrPeriodosPorFact
                Dim lstrPeriodoFacturado As String = ClsOrionCop.FstrPeriodoFinal(
                        ObjPeriodoIni_ItemProgStr.ObjValorPro, lshrPeriodosFact - 1)
                If lstrPeriodoFacturado < lstrPeriodoActual Then
                    Dim lstrPeriodoHoy = ClsOrionCop.FstrPeriodoDeFecha(Date.Today)
                    If lstrPeriodoHoy > lstrPeriodoActual Then
                        lblnPorFact = True
                    Else
                        Dim ldtmFechaFacturacionPerActual =
                                ObjServicio_ItemProgramaFact.DtmFechaFacturacionPeriodoActual
                        lblnPorFact = Day(ldtmFechaFacturacionPerActual) <= Day(Date.Today)
                    End If
                End If
            End If
        End If
        Return lblnPorFact
    End Function

    ''' <summary>
    ''' Actualiza el saldo por facturar cuando hay cambios en el periodo de inicio 
    ''' de facturación, o en la cantidad de periodos a facturar o en el valor a 
    ''' facturar en cada periodo
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub SActualiceSaldoActual()
        GobjPanDat.SControleProcesoObj(True)
        Dim ldecSaldoActual As Decimal
        If ObjPeriodoIni_ItemProgStr.BlnEsValido AndAlso ObjCantidadPeriodosShr.BlnEsValido AndAlso
                ObjValorPeriodo_ItemProgramaFactDec.BlnEsValido AndAlso
                ObjOrigen_ItemProgramaFacByt.BlnEsValido AndAlso
                ObjIdServicio_ItemProgramaFactShr.BlnEsValido Then
            If Not (ObjIdAno_ItemProgramaFactShr.ObjValorPro > 0 AndAlso
                    ObjOrigen_ItemProgramaFacByt.ObjValorPro =
                    EnuOrigenItemProgramaFactDef.EnuImportado) Then
                If ObjOrigen_ItemProgramaFacByt.ObjValorPro =
                        EnuOrigenItemProgramaFactDef.EnuAplicacion OrElse
                        ObjServicio_ItemProgramaFact.ObjGeneraProgramBln.ObjValorPro Then
                    Dim lentPeriodosFacturados As Integer
                    Dim lstrPeriodoInicio = ObjPeriodoIni_ItemProgStr.ObjValorPro
                    Dim lstrUltimoPeriodoFact = String.Empty
                    If ObjServicio_ItemProgramaFact.BlnEsCuotaAdministracion Then
                        If Not ObjServicio_ItemProgramaFact.ObjEsAjusteBln.ObjValorPro Then
                            Dim lobjAno As ClsAno = GobjParametros.ColAnos(ObjIdAno_ItemProgramaFactShr.ToString)
                            lstrUltimoPeriodoFact = lobjAno.FstrUltimoPerFacturado()
                        End If
                    Else
                        If ObjPeriodoIni_ItemProgStr.ObjValorPro <
                            GobjParametros.ObjAnoActual.StrIdPeriodoActual Then
                            If ObjIdPredio_ItemProgramaFactStr.ToString.Length > 0 Then
                                lstrUltimoPeriodoFact = FstrUltimoPeriodoFacturadoPredio()
                            ElseIf ObjIdCliente_ItemProgramaFactDbl.ObjValorPro <> 0 Then
                                lstrUltimoPeriodoFact = FstrUltimoPeriodoFacturadoCliente()
                            Else
                                Throw New ErrorInesperadoPanLException(
                                    "Item programa fact con problemas de destinatario!")
                            End If
                        ElseIf ObjPeriodoIni_ItemProgStr.ObjValorPro =
                            GobjParametros.ObjAnoActual.StrIdPeriodoActual Then
                            If ObjIdPredio_ItemProgramaFactStr.ToString.Length > 0 Then
                                lstrUltimoPeriodoFact = FstrUltimoPeriodoFacturadoPredio()
                            ElseIf ObjIdCliente_ItemProgramaFactDbl.ObjValorPro <> 0 Then
                                lstrUltimoPeriodoFact = FstrUltimoPeriodoFacturadoCliente()
                            Else
                                Throw New ErrorInesperadoPanLException(
                                    "Item programa fact con problemas de destinatario!")
                            End If
                            If lstrUltimoPeriodoFact <
                                GobjParametros.ObjAnoActual.StrIdPeriodoActual Then
                                lstrUltimoPeriodoFact = String.Empty
                            End If
                        Else
                            lstrUltimoPeriodoFact = String.Empty
                        End If
                    End If
                    If String.IsNullOrEmpty(lstrUltimoPeriodoFact) Then
                        lentPeriodosFacturados = 0
                    Else
                        lentPeriodosFacturados = ClsOrionCop.FentCantPeriodosEntrePeriodos(
                            lstrPeriodoInicio, lstrUltimoPeriodoFact) + 1
                        If lentPeriodosFacturados <= 0 Then
                            lentPeriodosFacturados = 0
                        End If
                    End If
                    Dim lentPeriodosPorFacturar As Integer = ObjCantidadPeriodosShr.ObjValorPro -
                        lentPeriodosFacturados
                    ldecSaldoActual = ObjValorPeriodo_ItemProgramaFactDec.ObjValorPro *
                            lentPeriodosPorFacturar
                Else
                    ldecSaldoActual = ObjValorPeriodo_ItemProgramaFactDec.ObjValorPro *
                        ObjCantidadPeriodosShr.ObjValorPro
                End If
                ObjSaldo_ItemProgramaFactDec.ObjValorPro = ldecSaldoActual
            End If
        End If
        GobjPanDat.SControleProcesoObj(False)
    End Sub

    Private Function FstrUltimoPeriodoFacturadoPredio() As String
        Dim lstrUltPerFac As String
        Dim lstrTabla = ClsItemFactura.SstrNombreTabla
        Dim lstrCamposSelect = {"MAX(" & ClsPeriodo_ItemFactStr.SstrNombreCampoBd & ")"}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd & " = " &
                ObjIdAno_ItemProgramaFactShr.ObjValorPro & " AND " &
                ClsIdServicio_ItemProgramaFactShr.SstrNombreCampoBd & " = " &
                ObjIdServicio_ItemProgramaFactShr.ToString & " AND " &
                ClsIdPredio_ItemProgramaFactStr.SstrNombreCampoBd & " = '" &
                ObjIdPredio_ItemProgramaFactStr.ObjValorPro & "'"
        Dim ldtbResul As DataTable = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamposSelect, {{}}, lstrFiltro)
        lstrUltPerFac = ClsPanorama.FobjValorCampo(ldtbResul(0)(0), EnuTipoValor.EnuString)
        If IsNothing(lstrUltPerFac) Then lstrUltPerFac = String.Empty
        Return lstrUltPerFac
    End Function

    Private Function FstrUltimoPeriodoFacturadoCliente() As String
        Dim lstrUltPerFac As String
        Dim lstrTablaPri = ClsFactura.SstrNombreTabla
        Dim lstrTablaSec = ClsItemFactura.SstrNombreTabla
        Dim lstrCamporPri = {ClsIdCliente_FactDbl.SstrNombreCampoBd}
        Dim lstrCamporRelPri = {StrCampoCarpeta, StrCampoCentroUtil,
                                ClsPrefijo_FactStr.SstrNombreCampoBd, ClsIdFacturaEnt.SstrNombreCampoBd}
        Dim lstrCamposSec = {"MAX(" & ClsPeriodo_ItemFactStr.SstrNombreCampoBd & ")"}
        Dim lstrCamposRelSec = {StrCampoCarpeta, StrCampoCentroUtil,
                                ClsPrefijo_ItemFactStr.SstrNombreCampoBd,
                                ClsIdFactura_ItemFactEnt.SstrNombreCampoBd}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion_Pri & " AND " &
                ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd & " = " &
                ObjIdAno_ItemProgramaFactShr.ObjValorPro & " AND " &
                ClsIdCliente_ItemProgramaFactDbl.SstrNombreCampoBd & " = " &
                ObjIdCliente_ItemProgramaFactDbl.ObjValorPro
        Dim ldtbResul As DataTable = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCamporPri,
                lstrTablaSec, lstrCamposSec, lstrCamporRelPri, lstrCamposRelSec, {{}}, lstrFiltro,
                Array.Empty(Of String), True)
        lstrUltPerFac = ClsPanorama.FobjValorCampo(ldtbResul(0)(1), EnuTipoValor.EnuString)
        If IsNothing(lstrUltPerFac) Then lstrUltPerFac = String.Empty
        Return lstrUltPerFac
    End Function

    Friend Sub SNumereObj()
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
            Dim lshrIdItemProgFac As Short
            GobjPanDat.SEjecuteSentenciaSql("FLUSH TABLES")
            lshrIdItemProgFac = ClsPanorama.FobjUltimaIdNumericaObjeto(SstrNombreTabla,
                    ClsOrdinal_ItemProgramaFact.SstrNombreCampoBd,
                    ObjOrdinal_ItemProgramaFact.EnuTipoValor, ClsOrionCop.StrFiltroUbicacion) + 1
            ObjOrdinal_ItemProgramaFact.ObjValorPro = lshrIdItemProgFac
        End If
    End Sub

    Private Sub SActualiceServiciosImportado(ashrIdAno As Short)
        Dim lobjAno As ClsAno, lstrKeySer As String
        Dim lobjServicio As ClsServicio
        Dim lstrPeriodoIni As String
        If ashrIdAno > 0 Then
            lstrPeriodoIni = ashrIdAno.ToString & "01"
            lobjAno = GobjParametros.ColAnos(ashrIdAno.ToString)
            For Each lobjServicio In lobjAno.ColServiciosAno
                If FblnSerImportado(lobjServicio) Then
                    SActualiceServicioImportado(lobjServicio, 12, lstrPeriodoIni)
                End If
            Next
            ClsCalculosServicios.SActualicePresAnoImportado(lobjAno)
        Else
            Dim lentCantPer As Integer
            For Each lstrKey As String In GarlServiciosImportados
                lstrKeySer = lstrKey.Split(",")(0) & "," & lstrKey.Split(",")(1)
                lentCantPer = CInt(lstrKey.Split(",")(2))
                lstrPeriodoIni = CInt(lstrKey.Split(",")(3))
                lobjServicio = GobjParametros.FobjServicio(lstrKeySer)
                SActualiceServicioImportado(lobjServicio, lentCantPer, lstrPeriodoIni)
                ClsCalculosServicios.SActualicePresModulosServicio(lobjServicio)
            Next
        End If
        GarlServiciosImportados = Nothing
        GobjParametros.SRefresqueObj()
    End Sub

    Private Sub SActualiceServicioImportado(aobjServicio As ClsServicio, aentCantPeriodos As Integer,
            astrPeriodoInicio As String)
        If aobjServicio.ObjEsAjusteBln.ObjValorPro Then
            aobjServicio.SMarqueAjustadoServicio()
        Else
            If Not aobjServicio.EnuPermisosObj And EnuPermisosDef.EnuModificar Then
                aobjServicio.EnuPermisosObj += EnuPermisosDef.EnuModificar
            End If
            aobjServicio.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
            If Not aobjServicio.ObjEsFactProgramableBln.ObjValorPro Then
                aobjServicio.ObjEsFactProgramableBln.ObjValorPro = True
            End If
            aobjServicio.ObjTipoBaseCalculoByt.ObjValorPro = EnuTipoBaseCalculo.EnuImportadas
            aobjServicio.ObjGeneraProgramBln.ObjValorPro = False
            aobjServicio.ObjEstaGenaradaProgramBln.ObjValorPro = True
            aobjServicio.ObjPeriodoInicioStr.ObjValorPro = astrPeriodoInicio
            aobjServicio.ObjCantPeriodos_ServicioShr.ObjValorPro = aentCantPeriodos
            If aobjServicio.ObjIdTipoServicioByt.ObjValorPro = EnuTipoServicio.EnuAnual Then
                aobjServicio.ObjEstaAjustadoBln.ObjValorPro =
                        Not aobjServicio.ObjMiAno.FblnHayFacsEnAno
            End If
            aobjServicio.SActualice(True)
        End If
    End Sub

    Friend Function FobjPredioImportando() As ClsPredio
        Dim lstrIdPredio = ObjIdPredio_ItemProgramaFactStr.ObjValorPro
        Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, lstrIdPredio}
        Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.EnuUnico)
        lobjPredio.SAbra(lobjValorLlave)
        Return lobjPredio
    End Function

    Private Function FblnEsValidoSer(ashrIdAno As Short, ashrIdServ As Short, aentIdReg As Integer,
            ByRef astrMens As String) As Boolean
        Dim lstrKey As String
        Dim lblnEsVali = ashrIdAno >= 0
        If Not lblnEsVali Then
            astrMens = "El Año en el Registro " & aentIdReg.ToString & " no es válido!"
            Return False
        End If
        lblnEsVali = ashrIdServ > 0
        If Not lblnEsVali Then
            astrMens = "El Servicio en el Registro " & aentIdReg.ToString & " no es válido!"
            Return False
        End If
        lstrKey = ashrIdAno.ToString & "," & ashrIdServ.ToString
        If ashrIdAno = 0 Then
            lblnEsVali = GobjParametros.ColServiciosPer.Contains(lstrKey)
        Else
            Dim lobjAno As ClsAno = GobjParametros.ColAnos(ashrIdAno.ToString())
            lblnEsVali = lobjAno.ColServiciosAno.Contains(lstrKey)
        End If
        If Not lblnEsVali Then
            astrMens = "El Servicio en el Registro " & aentIdReg.ToString &
                    " no existe en el año " & lstrKey.Substring(0, 4)
            Return False
        End If
        Return lblnEsVali
    End Function

    Private Function FblnEsDestiVali(adblIdTerc As Double, astrIdPredio As String,
            aentIdReg As Integer, ByRef astrMens As String) As Boolean
        Dim lblnEsVali = (adblIdTerc = 0 AndAlso
                Not String.IsNullOrEmpty(astrIdPredio)) OrElse
                (adblIdTerc > 0 AndAlso String.IsNullOrEmpty(astrIdPredio))
        If Not lblnEsVali Then
            astrMens = "En el Registro " & aentIdReg.ToString &
                    " se a ingresado Cliete y Predio lo cual no es válido!"
        End If
        Return lblnEsVali
    End Function

    Private Function FblnEsValidoConsumo(adrwOrigen As DataRow, ByRef adecValorPeriodo As Decimal,
            ByRef astrMens As String) As Boolean
        Const lstrLectAct = "lectura_actual"
        Const lstrLectAnt = "lectura_anterior"
        Const lstrVlrUni = "valor_unitario"
        Dim ldecLectActual As Decimal = ClsPanorama.FobjValorCampo(adrwOrigen(lstrLectAct),
                EnuTipoValor.EnuDecimal)
        Dim ldecLectAnt As Decimal = ClsPanorama.FobjValorCampo(adrwOrigen(lstrLectAnt),
                EnuTipoValor.EnuDecimal)
        Dim ldecVlrUnitario As Decimal = ClsPanorama.FobjValorCampo(adrwOrigen(lstrVlrUni),
                EnuTipoValor.EnuDecimal)
        Dim lblnEsValido = ldecLectActual = 0 AndAlso ldecLectAnt = 0
        If Not lblnEsValido Then
            lblnEsValido = ldecLectActual >= ldecLectAnt
            If Not lblnEsValido Then
                astrMens = " ,la lectura actual es inferior a la lectura anterior!"
            Else
                lblnEsValido = ldecLectActual > ldecLectAnt AndAlso ldecVlrUnitario > 0
                If Not lblnEsValido Then
                    astrMens = " ,el valor unitario debe ser mayor a cero!"
                End If
            End If
        End If
        If lblnEsValido Then
            adecValorPeriodo = Math.Round((ldecLectActual - ldecLectAnt) * ldecVlrUnitario)
        End If
        Return lblnEsValido
    End Function

    Friend Sub SCalculeVlrConsumo()
        Dim ldecValorCon As Decimal = (ObjLecturaActual_ItemProgramaFactDec.ObjValorPro -
                ObjLecturaAnterior_ItemProgramaFactDec.ObjValorPro) *
                ObjValorUnitario_ItemProgramaFactDec.ObjValorPro
        ObjCantidadPeriodosShr.ObjValorPro = 1
        ObjValorPeriodo_ItemProgramaFactDec.ObjValorPro = ldecValorCon
    End Sub

    Private Sub SAdecueValores()
        Dim lentUltPerFac As Integer, lentCantPeriPorFact As Integer
        If ObjServicio_ItemProgramaFact.ObjEsAjusteBln.ObjValorPro Then
            lentCantPeriPorFact = ObjCantidadPeriodosShr.ObjValorPro
        Else
            Dim lstrUltPeriodoFacturado = GobjParametros.ObjAnoActual.FstrUltimoPerFacturado
            If String.IsNullOrEmpty(lstrUltPeriodoFacturado) Then
                lentUltPerFac = 0
            Else
                lentUltPerFac = CInt(lstrUltPeriodoFacturado.Substring(4))
            End If
            lentCantPeriPorFact = 12 - lentUltPerFac
            ObjCantidadPeriodosShr.ObjValorPro = 12
        End If
        ObjSaldo_ItemProgramaFactDec.ObjValorPro = lentCantPeriPorFact *
                ObjValorPeriodo_ItemProgramaFactDec.ObjValorPro
    End Sub

    Private Function FblnSerImportado(aobjServicio As ClsServicio) As Boolean
        Dim lshrIdAno As Short = aobjServicio.ObjIdAno_ServicioShr.ObjValorPro
        Dim lshrIdSer As Short = aobjServicio.ObjIdServicioShr.ObjValorPro
        Dim lblnSiFueImp As Boolean
        For Each lstrServicio As String In GarlServiciosImportados
            lblnSiFueImp = lstrServicio.Split(",")(0) AndAlso lshrIdSer = lstrServicio.Split(",")(1)
            If lblnSiFueImp Then Exit For
        Next
        Return lblnSiFueImp
    End Function
#End Region
End Class

#Region "Clases de Propiedad"
Friend Class ClsCantidadPeriodosShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "CantidadPeriodos"
    Private ReadOnly MobjPadre As ClsItemProgramaFact = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "CantidadPeriodos"
        HStrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HenuTipoValor = EnuTipoValor.enuShort
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjValorIng As Object = HobjValorNew
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1,
                Short.MaxValue, HblnEsRequerido, EnuTipoValor)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                If HobjValorNew <> HobjValorOriginal Then
                    If MobjPadre.ObjServicio_ItemProgramaFact.ObjGeneraProgramBln.ObjValorPro AndAlso
                            MobjPadre.ObjServicio_ItemProgramaFact.BlnEsCuotaAdministracion Then
                        If Not (MobjPadre.BlnActualizando OrElse GblnImportando) AndAlso
                                MobjPadre.DecSaldo <> MobjPadre.DecValorTotal Then
                            HblnEsValido = False
                            HstrMens = "La Cantidad de Períodos de un Item de " &
                                    "Programación generado por el sistema, no puede ser modificado"
                        End If
                    End If
                End If
            End If
        Else
            HstrMens = "El valor ingresado no es válido!"
            If Not IsNothing(lobjValorIng) Then
                HstrMens = "El valor ingresado, " & lobjValorIng.ToString & ", no es válido!"
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
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsIdAno_ItemProgramaFactShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdAno"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdAñoItemProgramaFact"
        HenuTipoValor = EnuTipoValor.enuShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        Dim lblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, Year(Date.MaxValue),
                BlnEsRequerido, EnuTipoValor)
        If lblnEsValido Then
            If HobjValorNew > 0 Then
                lblnEsValido = GobjParametros.ColAnos.Contains(HobjValorNew.ToString)
            End If
            If lblnEsValido AndAlso HobjValorNew > 0 Then
                Dim lobjAno As ClsAno = GobjParametros.ColAnos(HobjValorNew.ToString)
                lblnEsValido = lobjAno.ColServiciosAno.Count > 0
                If Not lblnEsValido Then
                    HstrMens = "El año " & HobjValorNew.ToString & " no tiene servicios asociados!"
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
End Class

Friend Class ClsIdCliente_ItemProgramaFactDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTerceroCliente"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdCliente_ItemProgramaFact"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lblnEsValido As Boolean
        Dim lobjPadre As ClsItemProgramaFact = ObjPadre
        If lobjPadre.EnuTipoDeudor = EnuTipoDeudorDef.EnuCliente Then
            Dim lobjCliente As ClsCliente = lobjPadre.ObjPadre
            lblnEsValido = (HobjValorNew = lobjCliente.ObjIdClienteDbl.ObjValorPro)
        Else
            lblnEsValido = (IsNothing(HobjValorNew) OrElse String.IsNullOrEmpty(HobjValorNew) OrElse
                            HobjValorNew = 0)
            If lblnEsValido Then
                HobjValorNew = 0
            End If
        End If
        HblnEsValido = lblnEsValido
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        Dim lobjPadre As ClsItemProgramaFact = ObjPadre
        lobjPadre.ObjIdPredio_ItemProgramaFactStr.SValide()
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

Friend Class ClsIdPredio_ItemProgramaFactStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdPredio"
    Private ReadOnly MobjPadre As ClsItemProgramaFact = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdPredio_ItemProgramaFact"
        HenuTipoValor = EnuTipoValor.enuString
        HshrLongitud = 20
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        If MobjPadre.EnuTipoDeudor = EnuTipoDeudorDef.EnuPredio Then
            HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, HshrLongitud, True)
            If HblnEsValido Then
                Dim lobjPredio As ClsPredio
                If GblnImportando Then
                    lobjPredio = MobjPadre.FobjPredioImportando()
                Else
                    lobjPredio = MobjPadre.ObjPadre
                End If
                HblnEsValido = (HobjValorNew = lobjPredio.ObjIdPredioStr.ObjValorPro)
            End If
        Else
            HblnEsValido = (IsNothing(HobjValorNew) OrElse String.IsNullOrEmpty(HobjValorNew))
            If HblnEsValido Then
                HobjValorNew = Nothing
            End If
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        Dim lobjPadre As ClsItemProgramaFact = ObjPadre
        lobjPadre.ObjIdCliente_ItemProgramaFactDbl.SValide()
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

Friend Class ClsIdServicio_ItemProgramaFactShr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdServicio"
    Private ReadOnly MobjPadre As ClsItemProgramaFact = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = ObjPadre
        HstrNombre = "IdServicioItemProgramFact"
        HenuTipoValor = EnuTipoValor.enuUShort
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjPredio As ClsPredio
        Dim lobjCliente As ClsCliente = Nothing
        Dim lobjServicio As ClsServicio = Nothing
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Short.MaxValue,
                    BlnEsRequerido, EnuTipoValor)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                If MobjPadre.ObjIdAno_ItemProgramaFactShr.BlnEsValido Then
                    Dim lshrIdAno As Short = MobjPadre.ObjIdAno_ItemProgramaFactShr.ObjValorPro
                    Dim lstrKey = lshrIdAno.ToString & "," & HobjValorNew.ToString
                    Dim lcolItemsProgra As Collection
                    HstrMens = String.Empty
                    If Not GblnImportando Then
                        If TypeOf MobjPadre.ObjPadre Is ClsPredio Then
                            lobjPredio = MobjPadre.ObjPadre
                            lobjPredio.SRefresqueObj()
                            lcolItemsProgra = lobjPredio.ColItemsProgramaFact(lshrIdAno)
                            HblnEsValido = Not lcolItemsProgra.Contains(lstrKey)
                        Else
                            lobjCliente = MobjPadre.ObjPadre
                            lcolItemsProgra = lobjCliente.ColItemsProgramaFact
                            HblnEsValido = Not lcolItemsProgra.Contains(lstrKey)
                        End If
                    End If
                    If HblnEsValido Then
                        If lshrIdAno = 0 Then
                            HblnEsValido = GobjParametros.ColServiciosPer.Contains(lstrKey)
                            If HblnEsValido Then
                                lobjServicio = GobjParametros.ColServiciosPer(lstrKey)
                            End If
                        Else
                            Dim lobjAno As ClsAno = GobjParametros.ColAnos(lshrIdAno.ToString)
                            HblnEsValido = lobjAno.ColServiciosAno.Contains(lstrKey)
                            If HblnEsValido Then
                                lobjServicio = lobjAno.ColServiciosAno(lstrKey)
                            End If
                        End If
                        If HblnEsValido Then
                            If Not IsNothing(lobjCliente) Then
                                If Not IsNothing(lobjServicio) Then
                                    If lobjServicio.ObjFactAPropYPreAgrBln.ObjValorPro Then
                                        HblnEsValido = False
                                        HstrMens = "A un Cliente no se le pueden programar Servicios " &
                                                    "para ser facturados al Propietario!"
                                    End If
                                End If
                            End If
                        End If
                    Else
                        HstrMens = "El Servicio ya está programado!"
                    End If
                End If
            ElseIf MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                HblnEsValido = (HobjValorOriginal = HobjValorNew)
                If Not HblnEsValido Then
                    HstrMens = "No es permitido cambiar el Servicio!"
                End If
            End If
            If Not String.IsNullOrEmpty(HstrMens) Then
                SNotifiqueDatInv()
            End If
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        Dim lobjPadre As ClsItemProgramaFact = ObjPadre
        lobjPadre.ObjIdAno_ItemProgramaFactShr.SValide()
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

Friend Class ClsOrdinal_ItemProgramaFact
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Ordinal"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Ordinal"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsLlave = True
        HbytPosicionLlave = 2
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue,
                True, EnuTipoValor)
        If ObjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
            HstrMens = String.Empty
            If HblnEsValido Then
                If ObjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                    If Not BlnLeyendoOrigen Then
                        If HblnEsValido Then
                            Dim lobjLlavePrincipal() = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
                            If Not ObjPadre.FblnExisteLlave(lobjLlavePrincipal) Then
                                HstrMens = "El Ordinal del Item de la Programación para " &
                                        "Facturar ingresado no existe!"
                                HblnEsValido = False
                            End If
                        End If
                    End If
                ElseIf ObjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                    HblnEsValido = (HobjValorOriginal = HobjValorNew)
                    If Not HblnEsValido Then
                        HstrMens = "No es permitido cambiar la identidad a objeto alguno!"
                    End If
                End If
            Else
                HstrMens = "El Ordinal del Item de la Programación para Facturar ingresado " &
                        "no es válido!"
            End If
            If Not String.IsNullOrEmpty(HstrMens) Then
                SNotifiqueDatInv()
            End If
        Else
            HblnEsValido = True
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

Friend Class ClsOrigen_ItemProgramaFacByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Origen"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Origen Item Programacion"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuOrigenItemProgramaFactDef.EnuAplicacion,
                EnuOrigenItemProgramaFactDef.EnuImportado, True)
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

Friend Class ClsPeriodoIni_ItemProgStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PeriodoInicioFact"
    Private ReadOnly MobjPadre As ClsItemProgramaFact = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "PeriodoInicioFact"
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HshrLongitud = 6
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub

    Protected Overrides Sub SVaciePropiedad()
        HobjValorPro = GCSTRPERIODONULO
        HobjValorNew = HobjValorPro
        HblnEsValido = False
    End Sub

    Public Overrides Sub SValide()
        HblnEsValido = True
        If Not BlnLeyendoOrigen Then
            HblnEsValido = ClsPanorama.FblnEsValidoStringNumerico(HobjValorNew, ShrLongitud, ShrLongitud,
                        BlnEsRequerido)
            If HblnEsValido Then
                HstrMens = String.Empty
                If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
                    If Not IsNothing(MobjPadre.ObjServicio_ItemProgramaFact) Then
                        If Not MobjPadre.ObjServicio_ItemProgramaFact.ObjGeneraProgramBln.ObjValorPro Then
                            HblnEsValido = HobjValorNew >= GobjParametros.ObjAnoActual.StrIdPeriodoActual
                        End If
                    End If
                ElseIf MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando Then
                    If MobjPadre.ObjServicio_ItemProgramaFact.ObjGeneraProgramBln.ObjValorPro Then
                        If HobjValorNew < HobjValorOriginal Then
                            HblnEsValido = False
                            HstrMens = "El Período de Inicio para el cálculo es anterior " &
                                    "al Periodo de Inicio registrado!"
                        End If
                        If HobjValorOriginal <> HobjValorNew Then
                            HblnEsValido =
                                    Not MobjPadre.ObjServicio_ItemProgramaFact.ObjEstaGenaradaProgramBln.ObjValorPro
                            If Not HblnEsValido Then
                                HstrMens = "El Período de Inicio de un Item de " &
                                        "Programación generado por el sistema, " &
                                        "no puede ser modificado"
                            End If
                        End If
                    Else
                        HblnEsValido = (HobjValorNew >= GobjParametros.ObjAnoActual.StrIdPeriodoActual)
                        If Not HblnEsValido Then
                            HstrMens = "El Período de Inicio no puede ser anterior al " &
                                    "Período actual."
                        End If
                    End If
                End If
                If Not String.IsNullOrEmpty(HstrMens) Then
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
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsSaldo_ItemProgramaFactDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Saldo"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Saldo"
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HenuTipoValor = EnuTipoValor.enuDecimal
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0D,
                Decimal.MaxValue, True, EnuTipoValor)
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

Friend Class ClsValorPeriodo_ItemProgramaFactDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "ValorPeriodo"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "ValorPeriodo"
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HenuTipoValor = EnuTipoValor.enuDecimal
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = 0D
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0,
                Decimal.MaxValue, HblnEsRequerido, EnuTipoValor)
        If ObjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.EnuConsultando Then
            HstrMens = String.Empty
            If HblnEsValido Then
                HblnEsValido = HobjValorNew - Int(HobjValorNew) = 0
                If Not HblnEsValido Then
                    HstrMens = "El Valor ingresado debe ser sin Centavos!"
                End If
            Else
                HstrMens = "Elvalor ingresado no es valido. Debe ser un valor mayo oigual AddressOf cero"
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

Friend Class ClsConsumo_ItemProgramaFactDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Consumo"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Consumo_ItemProgramaFact"
        HenuTipoValor = EnuTipoValor.EnuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjValorIng As Object = HobjValorNew
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0,
                Decimal.MaxValue, False, EnuTipoValor)
        If Not HblnEsValido Then
            HstrMens = "El valor ingresado no es válido!"
            If Not IsNothing(lobjValorIng) Then
                HstrMens = "El valor ingresado, " & lobjValorIng.ToString & ", no es válido!"
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
            Return HobjValorPro.ToString
        End If
    End Function
End Class

Friend Class ClsValorUnitario_ItemProgramaFactDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "ValorUnitario"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "ValorUnitario_ItemProgramaFact"
        HenuTipoValor = EnuTipoValor.EnuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        Dim lobjValorIng = HobjValorNew
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, Decimal.MaxValue,
                BlnEsRequerido, EnuTipoValor)
        If Not HblnEsValido Then
            If Not IsNothing(lobjValorIng) Then
                HstrMens = "El valor ingresado, " & lobjValorIng.ToString & ", no es válido!"
            Else
                HstrMens = "El valor ingresado no es válido!"
            End If
        End If
        If Not String.IsNullOrEmpty(HstrMens) Then
            SNotifiqueDatInv()
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If HblnEsValido Then
            Dim lobjPadre As ClsItemProgramaFact = ObjPadre
            lobjPadre.SCalculeVlrConsumo()
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

Friend Class ClsLecturaActual_ItemProgramaFactDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "LecturaActual"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "LecturaActual_ItemProgramaFact"
        HenuTipoValor = EnuTipoValor.EnuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        Dim lobjValorIng As Object = HobjValorNew
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, Decimal.MaxValue,
                BlnEsRequerido, EnuTipoValor)
        If Not HblnEsValido Then
            If Not IsNothing(lobjValorIng) Then
                HstrMens = "El valor ingresado, " & lobjValorIng.ToString & ", no es válido!"
            Else
                HstrMens = "El valor ingresado no es válido!"
            End If
        Else
            Dim lobjPadre As ClsItemProgramaFact = ObjPadre
            HblnEsValido = HobjValorNew >=
                    lobjPadre.ObjLecturaAnterior_ItemProgramaFactDec.ObjValorPro
            If Not HblnEsValido Then
                HstrMens = "La lectura actual debe ser mayor o igual a la lectura anterior!"
            End If
            If HblnEsValido Then lobjPadre.SCalculeVlrConsumo()
        End If
        If Not String.IsNullOrEmpty(HstrMens) Then
            SNotifiqueDatInv()
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If HblnEsValido Then
            Dim lobjPadre As ClsItemProgramaFact = ObjPadre
            lobjPadre.SCalculeVlrConsumo()
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

Friend Class ClsLecturaAnterior_ItemProgramaFactDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "LecturaAnterior"
    Private ReadOnly MobjPadre As ClsItemProgramaFact = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "LecturaAnterior_ItemProgramaFact"
        HenuTipoValor = EnuTipoValor.EnuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        Dim lobjValorIng As Object = HobjValorNew
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, Decimal.MaxValue,
                BlnEsRequerido, EnuTipoValor)
        If Not HblnEsValido Then
            HstrMens = "El valor ingresado no es válido!"
            If Not IsNothing(lobjValorIng) Then
                HstrMens = "El valor ingresado, " & lobjValorIng.ToString & ", no es válido!"
            End If
        Else
            HblnEsValido = MobjPadre.ObjLecturaActual_ItemProgramaFactDec.ObjValorPro >=
                    HobjValorNew
            If Not HblnEsValido Then
                HstrMens = "La lectura actual debe ser mayor a la lectura anterior"
            End If
        End If
        If HblnEsValido Then
            Dim lobjPadre As ClsItemProgramaFact = ObjPadre
            HblnEsValido = HobjValorNew <=
                    lobjPadre.ObjLecturaAnterior_ItemProgramaFactDec.ObjValorPro
            If Not HblnEsValido Then
                HstrMens = "La lectura anterior debe ser menor o igual a la lectura actual!"
            End If
            If HblnEsValido Then lobjPadre.SCalculeVlrConsumo()

            lobjPadre.SCalculeVlrConsumo()
        End If
        If Not String.IsNullOrEmpty(HstrMens) Then
            SNotifiqueDatInv()
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If HblnEsValido Then
            Dim lobjPadre As ClsItemProgramaFact = ObjPadre
            lobjPadre.SCalculeVlrConsumo()
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
#End Region
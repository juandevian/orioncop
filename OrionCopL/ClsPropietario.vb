Public Class ClsPropietario
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriPropietarios"
    ' Variables
    Private MobjCliente As ClsCliente = Nothing
    Friend BlnExisteCliente As Boolean = True
#End Region
#Region "Constructores"
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
                   ClsIdCliente_PropDbl.SstrNombreCampoBd}
        Else
            HblnEsCreable = True
            HblnEsModificable = False
            HblnEsSuprimible = False
            HblnEsAnulable = False
            HenuTipoObjeto = EnuModoInstanciaObjDef.enuUnico
            lstrCamposSelect = {"*"}
        End If
        HcolTablas.Add(MCSTRNOMBRETABLA)
        HcolCamposSelect.Add(lstrCamposSelect)
    End Sub
    Friend Sub New(aobjPadre As ClsPredio, adrwObjeto As DataRow)
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
            Return EnuIdClasesPanDef.EnuPropietario
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Propietario"
        End Get
    End Property
    Friend Overrides ReadOnly Property HstrNombreObj As String
        Get
            Return Chr(34) & ObjCliente.ObjNombreCompletoStr.ObjValorPro & Chr(34)
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjIdCarpeta_PropShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_PropShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdCliente_PropDbl As New ClsIdCliente_PropDbl(Me)
    Friend ReadOnly Property ObjNombreCompleto_PropStr As New ClsNombreCompleto_PropStr(Me)
    Friend ReadOnly Property ObjPorcentajePartiDbl As New ClsPorcentajePartiDbl(Me)
    Friend ReadOnly Property ObjIdPredio_PropStr As New ClsIdPredio_PropStr(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjIdCarpeta_PropShr)
                HcolPropiedades.Add(ObjIdCentroUtil_PropShr)
                HcolPropiedades.Add(ObjIdCliente_PropDbl)
                HcolPropiedades.Add(ObjNombreCompleto_PropStr)
                HcolPropiedades.Add(ObjPorcentajePartiDbl)
                HcolPropiedades.Add(ObjIdPredio_PropStr)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        GobjPanDat.SControleProcesoObj(True)
        Dim lblnNoHayError = False
        Try
            If EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                GobjPanDat.SInicialiceTransaccion()
                SActuNombreComleto()
                If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                    Dim lobjPadre As ClsPredio = ObjPadre
                    ObjIdCarpeta_PropShr.ObjValorPro = GshrIdCarpeta
                    ObjIdCentroUtil_PropShr.ObjValorPro = GshrIdCentroUtil
                    If Not GblnImportando Then
                        Dim lobjPredio As ClsPredio = ObjPadre
                        ObjIdPredio_PropStr.ObjValorPro = lobjPredio.ObjIdPredioStr.ObjValorPro
                    End If
                    MyBase.SActualice(ablnExigeRequeridos)
                Else
                    MyBase.SActualice(ablnExigeRequeridos)
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
                GobjPanDat.SConfirmeTransaccion()
                GobjPanDat.SControleProcesoObj(False)
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    Protected Overrides Sub SVacie()
        MyBase.SVacie()
        BlnExisteCliente = True
    End Sub
#End Region
#Region "Cliente propietario"
    Friend ReadOnly Property ObjCliente As ClsCliente
        Get
            SAbraClientePropietario()
            Return MobjCliente
        End Get
    End Property
    Private Sub SAbraClientePropietario()
        Dim ldblIdCliente As Double = ObjIdCliente_PropDbl.ObjValorPro
        If MobjCliente Is Nothing OrElse ObjIdCliente_PropDbl.BlnCambio Then
            MobjCliente = New ClsCliente()
        End If
        If ldblIdCliente > 0 AndAlso ObjIdCliente_PropDbl.BlnEsValido Then
            If MobjCliente.ObjIdTerceroDbl.ObjValorPro <> ldblIdCliente Then
                MobjCliente.SAbra({GshrIdCarpeta, GshrIdCentroUtil, ldblIdCliente})
            End If
        End If
    End Sub
    Private Sub SActuNombreComleto()
        ObjNombreCompleto_PropStr.ObjValorPro =
                ObjCliente.ObjNombreCompletoStr.ObjValorPro
    End Sub
    Friend Overrides Function FblnSonValidosDatosOrigen(adtbOrigen As DataTable,
            astrColumnasRelacionadas As String(), ablnReinicie As Boolean,
            ByRef astrMens As String) As Boolean
        Dim lblnEsValido = False, i = 0, lstrColumnaOrigen As String
        Dim lbytIdCar As Byte, lbytIdCenutil As Byte, ldblIdTerCliente As Double
        Dim lstrIdPredio As String
        Dim lobjCliente As New ClsCliente(EnuModoInstanciaObjDef.EnuUnico)
        Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.EnuUnico)
        Dim lobjValorLlave As Object(), lstrFiltro As String, ldrwPropPred As DataRow()
        For Each ldrwOrigen As DataRow In adtbOrigen.Rows
            i += 1
            lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                    ClsIdCarpetaShr.SstrNombreCampoBd)
            lbytIdCar = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrColumnaOrigen),
                    EnuTipoValor.EnuString)
            lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                    ClsIdCentroUtilShr.SstrNombreCampoBd)
            lbytIdCenutil = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrColumnaOrigen),
                    EnuTipoValor.EnuString)
            lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                    ClsIdCliente_PropDbl.SstrNombreCampoBd)
            ldblIdTerCliente = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrColumnaOrigen),
                    EnuTipoValor.EnuDouble)
            lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                    ClsIdPredio_PropStr.SstrNombreCampoBd)
            lstrIdPredio = ClsPanorama.FobjValorCampo(ldrwOrigen(lstrColumnaOrigen),
                    EnuTipoValor.EnuString)
            ' Validar carpeta
            lblnEsValido = lbytIdCar = GshrIdCarpeta
            If Not lblnEsValido Then
                astrMens = "La carpeta del registro " & i.ToString & " no es la actual!"
                Exit For
            End If
            ' Validar Centro Utilidad
            lblnEsValido = lbytIdCenutil = GshrIdCentroUtil
            If Not lblnEsValido Then
                astrMens = "La copropiedad del registro " & i.ToString &
                        " no es la actual!"
                Exit For
            End If
            ' Validar Cliente
            lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, ldblIdTerCliente}
            lobjCliente.SAbra(lobjValorLlave)
            lblnEsValido = lobjCliente.BlnExiste
            If Not lblnEsValido Then
                astrMens = "El cliente del registro " & i.ToString & " no existe!"
                Exit For
            End If
            ' Validar predio
            lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrIdPredio}
            lobjPredio.SAbra(lobjValorLlave)
            lblnEsValido = lobjPredio.BlnExiste
            If Not lblnEsValido Then
                astrMens = "El predio del registro " & i.ToString & " no existe!"
                Exit For
            End If
            lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                    ClsIdPredio_PropStr.SstrNombreCampoBd)
            lstrFiltro = lstrColumnaOrigen & " = '" & lstrIdPredio & "'"
            ldrwPropPred = adtbOrigen.Select(lstrFiltro)
            lstrColumnaOrigen = FstrColumnaOrigen(astrColumnasRelacionadas,
                    ClsPorcentajePartiDbl.SstrNombreCampoBd)
            lblnEsValido = FblnEsValidoPorcientoParticip(ldrwPropPred, lstrColumnaOrigen)
            If Not lblnEsValido Then
                astrMens = "La propiedad del predio '" & lstrIdPredio & "' no es del 100%!"
                Exit For
            End If
        Next
        Return lblnEsValido
    End Function
    Private Function FblnEsValidoPorcientoParticip(adrwsPredprop As DataRow(),
            astrNombreColumna As String) As Boolean
        Dim ldblPorParticipa As Double, ldbllTotPar = 0.0
        For Each ldrwPrePro As DataRow In adrwsPredprop
            ldblPorParticipa = ClsPanorama.FobjValorCampo(ldrwPrePro(
                    astrNombreColumna), EnuTipoValor.EnuDouble)
            ldbllTotPar += ldblPorParticipa
        Next
        Return ldbllTotPar = 1
    End Function
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsIdCliente_PropDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTerceroPropietario"
    Dim MstrNombreCli As String = String.Empty
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdPropietario"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 3
    End Sub

    Public Overrides Sub SValide()
        HstrMens = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCDBLMINTERC,
                GCDBLMAXTERC, BlnEsRequerido)
        Dim lobjPadre As ClsPropietario = ObjPadre
        Dim lobjPredio As ClsPredio = lobjPadre.ObjPadre
        If HblnEsValido Then
            If Not BlnLeyendoOrigen Then
                If ObjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
                    Dim lobjCliente As New ClsCliente(EnuModoInstanciaObjDef.EnuUnico)
                    Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
                    lobjCliente.SAbra(lobjValorLlave)
                    lobjPadre.BlnExisteCliente = lobjCliente.BlnExiste
                    If Not lobjCliente.BlnExiste Then
                        HstrMens = "Un Propietario con el número de identificación ingresado, '" &
                                HobjValorNew.ToString & "', no existe!"
                        HblnEsValido = False
                    Else
                        MstrNombreCli = lobjCliente.ObjNombreCompletoStr.ObjValorPro
                    End If
                    If Not GblnImportando Then
                        If HblnEsValido AndAlso lobjPredio.ColPropietarios.Contains(HobjValorNew) Then
                            HstrMens = "El predio ya tiene este cliente asignado como propietario!"
                            HblnEsValido = False
                        End If
                    End If
                End If
            End If
        Else
            If (ObjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando AndAlso
                        HobjValorNew <> 0) OrElse ObjPadre.EnuEstadoActualizacion =
                        EnuEstadoObjetoDef.enuConsultando Then
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

    Friend ReadOnly Property StrNombreCli As String
        Get
            If BlnEsValido Then
                Return MstrNombreCli
            Else
                Return String.Empty
            End If
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

Friend Class ClsIdPredio_PropStr
    'Herencia
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdPredio"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Predio de propietario"
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
        If Not HblnEsValido Then
            If ObjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando AndAlso
                    Not String.IsNullOrEmpty(HobjValorNew) Then
                HstrMens = "El valor ingresado, '" & HobjValorNew.ToString &
                        "', no es valido!"
            End If
        End If
        If Not String.IsNullOrEmpty(HstrMens) Then
            Dim lobjPredio As ClsPredio = ObjPadre.ObjPadre
            lobjPredio.SLevanteEventoNot(Me, HstrMens, "", 0, EnuSeveridadNot.EnuInformacion)
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

Friend Class ClsNombreCompleto_PropStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "NombreCompleto"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "NombreCompleto"
        HshrLongitud = 100
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
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

Friend Class ClsPorcentajePartiDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PorcentajeParticipacion"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Proporción de propiedad"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnRegistrarLogCambio = True
    End Sub

    Public Overrides Sub SValide()
        Dim lobjPadre As ClsPropietario = ObjPadre
        Dim lobjPredio As ClsPredio = lobjPadre.ObjPadre
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0,
                1, BlnEsRequerido, HenuTipoValor)
        HblnEsValido = HblnEsValido AndAlso HobjValorNew > 0
        If Not HblnEsValido AndAlso lobjPadre.ObjIdCliente_PropDbl.BlnEsValido Then
            lobjPredio.SLevanteEventoNot(Me,
                    "El porcentaje debe ser un Número mayor a cero y menor o igual a cien!", "", 0,
                    EnuSeveridadNot.EnuInformacion)
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
            Return Format(HobjValorPro, "p")
        End If
    End Function
End Class
#End Region

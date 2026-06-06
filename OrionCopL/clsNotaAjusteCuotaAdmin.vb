Friend Class ClsNotaAjusteCuotaAdmin
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriNotasAjusteCuota"
    '
    Private MobjPredioAgrNAjuste As ClsPredio = Nothing
    Private MobjClienteNota As clsCliente = Nothing
    Private MdtbNovedadesAnt As DataTable = Nothing
    Private McolNovedadesAnt As Collection = Nothing
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia un objeto Nota Ajuste en modo único
    ''' </summary>
    Public Sub New()
        HobjPadre = Nothing
        HblnEsCreable = False
        HblnEsModificable = False
        HblnEsAnulable = False
        HblnEsSuprimible = False
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuUnico
        HcolTablas.Add(MCSTRNOMBRETABLA)
        HcolCamposSelect.Add({"*"})
    End Sub
    ''' <summary>
    ''' Instancia un objeto Nota Ajuste en modo navegable
    ''' </summary>
    Public Sub New(astrPref As String)
        HobjPadre = Nothing
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion
        lstrFiltro &= " AND " & ClsPrefijo_NotaAjusteStr.SstrNombreCampoBd & " = '" & astrPref & "'"
        HcolFiltros.Add(lstrFiltro)
        Dim lstrCamposSelect = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsPrefijo_NotaAjusteStr.SstrNombreCampoBd, ClsIdNotaAjusteEnt.SstrNombreCampoBd}
        HblnEsCreable = False
        HblnEsSuprimible = False
        HblnEsModificable = False
        HblnEsAnulable = False
        HcolTablas.Add(MCSTRNOMBRETABLA)
        HcolCamposSelect.Add(lstrCamposSelect)
    End Sub
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Administrador al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwObjeto">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Public Sub New(aobjPadre As clsCliente, adrwObjeto As DataRow)
        HobjPadre = aobjPadre
        henuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        hblnEsAnulable = False
        HblnEsSuprimible = False
        hblnEsModificable = False
        '
        drwRegistroActual = adrwObjeto
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
            Return EnuIdClasesPanDef.enuNotaAjusteCuotaAdmin
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Nota Ajuste Cuota Administración"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjFecha_NotaAjusteDtm As New ClsFecha_NotaAjusteDtm(Me)
    Friend ReadOnly Property ObjIdAnticipo_NotaAjusteEnt As New ClsIdAnticipo_NotaAjusteEnt(Me)
    Friend ReadOnly Property ObjIdCarpeta_NotaAjusteShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_NotaAjusteShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdCliente_NotaAjusteDbl As New ClsIdCliente_NotaAjusteDbl(Me)
    Friend ReadOnly Property ObjIdNotaAjusteEnt As New ClsIdNotaAjusteEnt(Me)
    Friend ReadOnly Property ObjIdPredio_NotaAjusteStr As New ClsIdPredio_NotaAjusteStr(Me)
    Friend ReadOnly Property ObjIdPredioAgrupador_NotaAjusteStr As New ClsIdPredioAgrupador_NotaAjusteStr(Me)
    Friend ReadOnly Property ObjIdUsuario_NotaAjusteStr As New ClsIdUsuarioStr(Me)
    Friend ReadOnly Property ObjPrefijo_NotaAjusteStr As New ClsPrefijo_NotaAjusteStr(Me)
    Friend ReadOnly Property ObjValor_NotaAjusteDec As New ClsValor_NotaAjusteDec(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjFechaCreacionDtm)
                HcolPropiedades.Add(ObjOrigenInstanciaStr)
                HcolPropiedades.Add(ObjFecha_NotaAjusteDtm)
                HcolPropiedades.Add(ObjIdAnticipo_NotaAjusteEnt)
                HcolPropiedades.Add(ObjIdCarpeta_NotaAjusteShr)
                HcolPropiedades.Add(ObjIdCentroUtil_NotaAjusteShr)
                HcolPropiedades.Add(ObjIdCliente_NotaAjusteDbl)
                HcolPropiedades.Add(ObjIdNotaAjusteEnt)
                HcolPropiedades.Add(ObjIdPredio_NotaAjusteStr)
                HcolPropiedades.Add(ObjIdPredioAgrupador_NotaAjusteStr)
                HcolPropiedades.Add(ObjIdUsuario_NotaAjusteStr)
                HcolPropiedades.Add(ObjPrefijo_NotaAjusteStr)
                HcolPropiedades.Add(ObjValor_NotaAjusteDec)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    Friend Property StrIdCuentaDb As String = String.Empty
    Friend Property StrServicios As String = String.Empty
    Friend Property ObjClienteNota As ClsCliente
        Get
            Dim lobjValorLlave As Object() = {ObjIdCarpeta_NotaAjusteShr.ObjValorPro,
                ObjIdCentroUtil_NotaAjusteShr.ObjValorPro, ObjIdCliente_NotaAjusteDbl.ObjValorPro}
            MobjClienteNota = New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
            MobjClienteNota.SAbra(lobjValorLlave)
            Return MobjClienteNota
        End Get
        Set(value As ClsCliente)
            MobjClienteNota = value
            HobjPadre = MobjClienteNota
        End Set
    End Property
    ''' <summary>
    ''' Devuelve uns string compuesto por el prefijo de la nota y el id de la nota separados por un
    ''' guion. Si no existe el prefijo devuelve solo el id de la factura
    ''' </summary>
    ''' <value></value>
    Friend ReadOnly Property StrNumeroNotaAjuste As String
        Get
            Dim lstrNumeroNotaCon As String = ClsPanorama.FstrNumeroDcto(ObjPrefijo_NotaAjusteStr.ObjValorPro,
                    ObjIdNotaAjusteEnt.ObjValorPro)
            Return lstrNumeroNotaCon
        End Get
    End Property
    Friend ReadOnly Property ObjPredioAgrNAjuste As ClsPredio
        Get
            If IsNothing(MobjPredioAgrNAjuste) Then
                If Not String.IsNullOrEmpty(ObjIdPredioAgrupador_NotaAjusteStr.ToString) AndAlso
                        ObjIdPredioAgrupador_NotaAjusteStr.ToString <> GCSTRSINPA Then
                    MobjPredioAgrNAjuste = New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
                    Dim lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil,
                            ObjIdPredioAgrupador_NotaAjusteStr.ObjValorPro}
                    MobjPredioAgrNAjuste.SAbra(lobjValorLlave)
                    If Not MobjPredioAgrNAjuste.BlnExiste Then
                        MobjPredioAgrNAjuste = Nothing
                    End If
                End If
            End If
            Return MobjPredioAgrNAjuste
        End Get
    End Property
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjIdNotaAjusteEnt.ToString
        End Get
    End Property
#End Region
#Region "Anticipo"
    Friend ReadOnly Property ObjAnticipo As clsAnticipo
        Get
            Dim lobjAnticipo As ClsAnticipo = ObjIdAnticipo_NotaAjusteEnt.ObjAnticipo
            Return lobjAnticipo
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SInicialiceObj()
        ObjFechaCreacionDtm.ObjValorPro = Date.Now
        ObjOrigenInstanciaStr.ObjValorPro = GstrOrigenActual
        ObjIdCarpeta_NotaAjusteShr.ObjValorPro = GshrIdCarpeta
        ObjIdCentroUtil_NotaAjusteShr.ObjValorPro = GshrIdCentroUtil
        ObjIdUsuario_NotaAjusteStr.ObjValorPro = GstrIdUsuario
        ObjPrefijo_NotaAjusteStr.ObjValorPro = GobjParametros.
                FstrPrefijoDoc(EnuTipoDocOri.EnuNotaAjuste)
    End Sub
    Protected Overrides Sub SVacie()
        MyBase.SVacie()
        MobjClienteNota = Nothing
        MdtbNovedadesAnt = Nothing
        McolNovedadesAnt = Nothing
    End Sub
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        Dim lblnNoHayError = False
        gobjPanDat.sControleProcesoObj(True)
        Try
            gobjPanDat.sInicialiceTransaccion()
            If enuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                sNumereObj()
            End If
            If Not gblnActualizandoApp Then
                sGenereAnticipoAjuste()
            End If
            MyBase.sActualice(ablnExigeRequeridos)
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
                gobjPanDat.sConfirmeTransaccion()
                gobjPanDat.sControleProcesoObj(False)
            Else
                gobjPanDat.sAborteTransaccion()
                gobjPanDat.sControleProcesoObj(False, True)
            End If
        End Try
    End Sub
#End Region
#Region "Procedimientos del objeto"
    Private Sub SNumereObj()
        If enuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            Dim lentIdNotaAjuste As Integer
            lentIdNotaAjuste = clsPanorama.FobjUltimaIdNumericaObjeto(sstrNombreTabla,
                    ObjIdNotaAjusteEnt.strNombreCampoBD, ObjIdNotaAjusteEnt.enuTipoValor,
                    clsOrionCop.strFiltroUbicacion) + 1
            ObjIdNotaAjusteEnt.ObjValorPro = lentIdNotaAjuste
        End If
    End Sub
    Friend Sub SGenereAnticipoAjuste()
        Dim lobjAnticipo As ClsAnticipo = ObjClienteNota.FobjNuevoAnticipo(
                ObjFecha_NotaAjusteDtm.ObjValorPro, ObjIdPredioAgrupador_NotaAjusteStr.ObjValorPro)
        With lobjAnticipo
            .ObjIdTipoDocOrigen_AntByt.ObjValorPro = EnuTipoDocOri.EnuNotaAjuste
            .ObjIdPredio_AntStr.ObjValorPro = ObjIdPredio_NotaAjusteStr.ObjPredio.ObjIdPredioStr.ObjValorPro
            .ObjServicios_AntStr.ObjValorPro = StrServicios
            .ObjIdDocOrigen_AntEnt.ObjValorPro = ObjIdNotaAjusteEnt.ObjValorPro
            .ObjPrefijoDocOrigen_AntStr.ObjValorPro = ObjPrefijo_NotaAjusteStr.ObjValorPro
            .ObjFechaAnticipoDtm.ObjValorPro = ObjFecha_NotaAjusteDtm.ObjValorPro
            .SGenereNovedadAntRecibido(ObjFecha_NotaAjusteDtm.ObjValorPro, StrIdCuentaDb,
                    ObjValor_NotaAjusteDec.ObjValorPro)
            .SActualice(True)
            ObjIdAnticipo_NotaAjusteEnt.ObjValorPro = lobjAnticipo.ObjIdAnticipoEnt.ObjValorPro
        End With
    End Sub
    Friend Function FstrAliasCon() As String
        Dim lstrAliasCon = String.Empty
        If ObjPredioAgrNAjuste IsNot Nothing Then
            lstrAliasCon = ObjPredioAgrNAjuste.ObjAliasContStr.ToString
        End If
        If String.IsNullOrEmpty(lstrAliasCon) Then
            lstrAliasCon = ObjIdCliente_NotaAjusteDbl.ToString
        End If
        Return lstrAliasCon
    End Function
#End Region
#Region "Novedades"
    Friend ReadOnly Property ColNovedadesAnt As Collection
        Get
            If IsNothing(McolNovedadesAnt) Then
                McolNovedadesAnt = New Collection
                sCargueDtbNovedadesNota()
                If Not IsNothing(mdtbNovedadesAnt) AndAlso mdtbNovedadesAnt.Rows.Count > 0 Then
                    Dim ldrwNovedades() As DataRow = mdtbNovedadesAnt.Select
                    For Each ldrwNovedad As DataRow In ldrwNovedades
                        Dim lobjNovedad As New clsNovedad(Me, ldrwNovedad)
                        lobjNovedad.sLeaValores(True)
                        McolNovedadesAnt.Add(lobjNovedad, lobjNovedad.ObjIdNovedadShr.ToString)
                    Next
                End If
            End If
            Return McolNovedadesAnt
        End Get
    End Property
    Friend ReadOnly Property DtbNovedadesNota As DataTable
        Get
            SCargueDtbNovedadesNota()
            Return MdtbNovedadesAnt
        End Get
    End Property
    Private Sub SCargueDtbNovedadesNota()
        If IsNothing(MdtbNovedadesAnt) Then
            Dim lstrIdNota = ObjIdNotaAjusteEnt.ToString
            If String.IsNullOrEmpty(lstrIdNota) Then lstrIdNota = "0"
            Dim lstrTabla = ClsNovedadAnticipo.SstrNombreTabla
            Dim lstrCamposSelect = {"*"}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsIdTipoDocOrigen_NovAntByt.SstrNombreCampoBd &
                    " = " & EnuTipoDocOri.EnuNotaAjuste & " AND " &
                    ClsPrefijoDocOrigen_NovAntStr.SstrNombreCampoBd & " = ''" &
                    " AND " & ClsIdDocOrigen_NovAntEnt.SstrNombreCampoBd & " = " & lstrIdNota

            Dim lstrIndice = {{ClsPrefijoDocOrigen_NovAntStr.SstrNombreCampoBd, "ASC"},
                              {ClsIdDocOrigen_NovAntEnt.SstrNombreCampoBd, "ASC"},
                              {ClsIdNovedadShr.SstrNombreCampoBd, "ASC"}}
            MdtbNovedadesAnt = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamposSelect, lstrIndice, lstrFiltro)
        End If
    End Sub
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsFecha_NotaAjusteDtm
    Inherits clsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaNota"
    Private ReadOnly MobjPadre As clsNotaAjusteCuotaAdmin = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Fecha Ajuste Cuota Admin"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim ldtmFechaMin = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
        Dim ldtmFechaMax = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
        If GblnActualizandoApp Then
            ldtmFechaMin = DateSerial(2016, 1, 1)
        End If
        If ldtmFechaMax > Date.Today Then
            ldtmFechaMax = Now
        End If
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax, BlnEsRequerido)
        Else
            HblnEsValido = (HobjValorNew = HobjValorOriginal)
        End If
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(HobjValorPro) Then
            Return Format(HobjValorPro, GCSTRFMTFECHASIMPLE)
        Else
            Return Format(GCDTMFECHANULA, GCSTRFMTFECHASIMPLE)
        End If
    End Function
End Class
Friend Class ClsIdAnticipo_NotaAjusteEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdAnticipo"
    Private ReadOnly MobjPadre As ClsNotaAjusteCuotaAdmin = Nothing
    Private MobjAnticipo As ClsAnticipo = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdAnticipo Nota Ajuste"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MyBase.SVaciePropiedad()
        MobjAnticipo = Nothing
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue, BlnEsRequerido,
                EnuTipoValor.enuInteger)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            Else
                MobjAnticipo = ObjAnticipo
                HblnEsValido = MobjAnticipo.BlnExiste AndAlso
                        MobjPadre.ObjIdPredioAgrupador_NotaAjusteStr.ObjValorPro =
                        MobjAnticipo.ObjIdPredioAgrupador_AntStr.ObjValorPro
            End If
        End If
    End Sub
    Friend ReadOnly Property ObjAnticipo() As ClsAnticipo
        Get
            If IsNothing(MobjAnticipo) Then
                MobjAnticipo = New ClsAnticipo(EnuModoInstanciaObjDef.enuUnico)
            End If
            Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
            MobjAnticipo.SAbra(lobjValorLlave)
            Return MobjAnticipo
        End Get
    End Property
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        MobjPadre.ObjIdPredioAgrupador_NotaAjusteStr.SValide()
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
Friend Class ClsIdCliente_NotaAjusteDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTerceroCliente"
    Private ReadOnly MobjPadre As ClsNotaAjusteCuotaAdmin = Nothing
    Private MobjCliente As ClsCliente = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Id Tercero Cliente NotaAjuste"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjValorIng = HobjValorNew
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCDBLMINTERC, GCDBLMAXTERC,
                BlnEsRequerido)
        HstrMens = String.Empty
        If HblnEsValido Then
            Dim lobjLlavePrincipal() = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
            MobjCliente = New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
            MobjCliente.SAbra(lobjLlavePrincipal)
            If Not MobjCliente.BlnExiste Then
                HstrMens = "El Cliente ingresado no es existe!"
            End If
            MobjPadre.ObjClienteNota = MobjCliente
        ElseIf Not String.IsNullOrEmpty(lobjValorIng.ToString) Then
            HstrMens = "La Id. del Cliente ingresada, '" & lobjValorIng.ToString & ",  no es válida!"
        Else
            HobjValorNew = ""
        End If
        If Not String.IsNullOrEmpty(HstrMens) Then
            SNotifiqueDatInv()
        End If
    End Sub
    Friend ReadOnly Property StrNombreCliente As String
        Get
            If Not IsNothing(MobjCliente) AndAlso MobjCliente.BlnExiste Then
                Return MobjCliente.ObjNombreCompletoStr.ObjValorPro
            Else
                Return ""
            End If
        End Get
    End Property
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
Friend Class ClsIdNotaAjusteEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdNotaAjusteCuotaAdmin"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Id Nota Ajuste Cuota Admin."
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 3
    End Sub
    Public Overrides Sub SValide()
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
End Class
Friend Class ClsIdPredio_NotaAjusteStr
    'Herencia
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdPredio"
    Private ReadOnly MobjPadre As ClsNotaAjusteCuotaAdmin = Nothing
    Private MobjPredio As ClsPredio = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdPredio Nota Ajuste"
        HenuTipoValor = EnuTipoValor.enuString
        HshrLongitud = 20
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MyBase.SVaciePropiedad()
        MobjPredio = Nothing
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            MobjPredio = New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
            Dim lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
            MobjPredio.SAbra(lobjValorLlave)
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = MobjPredio.BlnExiste
            Else
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            End If
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        MobjPadre.ObjIdPredioAgrupador_NotaAjusteStr.SValide()
    End Sub
    Friend ReadOnly Property ObjPredio As ClsPredio
        Get
            Return MobjPredio
        End Get
    End Property
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
Friend Class ClsIdPredioAgrupador_NotaAjusteStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdPredioAgrupador"
    Private ReadOnly MobjPadre As ClsNotaAjusteCuotaAdmin = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Id PredioAgrupador Nota Ajuste"
        HenuTipoValor = EnuTipoValor.enuString
        HshrLongitud = 20
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud,
                BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = (MobjPadre.ObjIdPredio_NotaAjusteStr.ObjPredio.ObjIdPredioAgrupadorStr.
                        ObjValorPro = HobjValorNew)
            Else
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            End If
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        MobjPadre.ObjIdAnticipo_NotaAjusteEnt.SValide()
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(ObjValorPro) Then
            Dim lstrIdPredioAgr = GCSTRSINPA
            If HobjValorPro <> "" Then
                lstrIdPredioAgr = HobjValorPro
            End If
            Return lstrIdPredioAgr
        Else
            Return ""
        End If
    End Function
End Class
Friend Class ClsPrefijo_NotaAjusteStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PrefijoNota"
    Private ReadOnly MobjPadre As ClsCBObjetoPan = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Prefijo Nota Ajuste"
        HshrLongitud = 5
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsLlave = True
        HbytPosicionLlave = 2
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then
            HobjValorNew = String.Empty
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = (HobjValorNew = GobjParametros.
                                FstrPrefijoDoc(EnuTipoDocOri.EnuNotaAjuste))
            Else
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
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
Friend Class ClsValor_NotaAjusteDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Valor"
    Private ReadOnly MobjPadre As ClsNotaAjusteCuotaAdmin = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Valor"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0.01,
                Decimal.MaxValue, BlnEsRequerido, HenuTipoValor)
        If Not HblnEsValido Then
            If HobjValorNew = 0 Then
                HblnEsValido = MobjPadre.ObjAnuladoBln.ObjValorPro
            End If
        Else
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
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
            Return Format(HobjValorPro, "c")
        End If
    End Function
End Class
#End Region
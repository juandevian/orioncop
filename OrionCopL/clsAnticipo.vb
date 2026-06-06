Friend Class ClsAnticipo
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriAnticipos"
    ' Variables de modulo
    Private MobjPredioAgrAnt As ClsPredio = Nothing
    Private McolNovedadesAnt As Collection = Nothing
    Private McolNovedadesRev As Collection = Nothing
    Private MdtbNovedadesAnt As DataTable = Nothing
    Private MdtbNotasCon As DataTable = Nothing
    Private McolNotasCon As Collection = Nothing
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
    Public Sub New(aenuModoInstanciaObj As enuModoInstanciaObjDef)
        If aenuModoInstanciaObj = enuModoInstanciaObjDef.enuDeColeccion Then
            Throw New ErrorInesperadoPanLException("Con este Constructor no se puede instanciar un Objeto de Colección!")
        End If
        Dim lstrCamposSelect As String()
        HobjPadre = Nothing
        hblnEsCreable = False
        HblnEsModificable = False
        HblnEsSuprimible = False
        hblnEsAnulable = False
        If aenuModoInstanciaObj = enuModoInstanciaObjDef.enuNavegable Then
            lstrCamposSelect = {StrCampoCarpeta, StrCampoCentroUtil,
                                ClsIdAnticipoEnt.SstrNombreCampoBd}
            HcolFiltros.Add(ClsOrionCop.StrFiltroUbicacion)
        Else
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
    ''' <param name="adrwAnticipo">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Public Sub New(aobjPadre As ClsCliente, adrwAnticipo As DataRow)
        HobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        '
        DrwRegistroActual = adrwAnticipo
        DtbTablaColeccion = DrwRegistroActual.Table
        HblnEsAnulable = False
        HblnEsSuprimible = False
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
            Return EnuIdClasesPanDef.EnuAnticipo
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Anticipo"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjCreditos_AntDec As New ClsCreditos_AntDec(Me)
    Friend ReadOnly Property ObjDebitos_AntDec As New ClsDebitos_AntDec(Me)
    Friend ReadOnly Property ObjFechaAnticipoDtm As New ClsFechaAnticipoDtm(Me)
    Friend ReadOnly Property ObjIdAnticipoEnt As New ClsIdAnticipoEnt(Me)
    Friend ReadOnly Property ObjIdCarpeta_AntShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_AntShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdCliente_AntDbl As New ClsIdCliente_AntDbl(Me)
    Friend ReadOnly Property ObjIdDocOrigen_AntEnt As New ClsIdDocOrigen_AntEnt(Me)
    Friend ReadOnly Property ObjIdPredio_AntStr As New ClsIdPredio_AntStr(Me)
    Friend ReadOnly Property ObjIdPredioAgrupador_AntStr As New ClsIdPredioAgrupador_AntStr(Me)
    Friend ReadOnly Property ObjIdTipoDocOrigen_AntByt As New ClsIdTipoDocOrigen_AntByt(Me)
    Friend ReadOnly Property ObjPrefijoDocOrigen_AntStr As New ClsPrefijoDocOrigen_AntStr(Me)
    Friend ReadOnly Property ObjServicios_AntStr As New ClsServicios_AntStr(Me)
    Friend ReadOnly Property ObjValor_AntDec As New ClsValor_AntDec(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjAnuladoBln)
                HcolPropiedades.Add(ObjFechaCreacionDtm)
                HcolPropiedades.Add(ObjCreditos_AntDec)
                HcolPropiedades.Add(ObjDebitos_AntDec)
                HcolPropiedades.Add(ObjFechaAnticipoDtm)
                HcolPropiedades.Add(ObjIdAnticipoEnt)
                HcolPropiedades.Add(ObjIdCarpeta_AntShr)
                HcolPropiedades.Add(ObjIdCentroUtil_AntShr)
                HcolPropiedades.Add(ObjIdCliente_AntDbl)
                HcolPropiedades.Add(ObjIdPredio_AntStr)
                HcolPropiedades.Add(ObjIdPredioAgrupador_AntStr)
                HcolPropiedades.Add(ObjIdDocOrigen_AntEnt)
                HcolPropiedades.Add(ObjIdTipoDocOrigen_AntByt)
                HcolPropiedades.Add(ObjPrefijoDocOrigen_AntStr)
                HcolPropiedades.Add(ObjServicios_AntStr)
                HcolPropiedades.Add(ObjValor_AntDec)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    Friend ReadOnly Property DecAnticipoPorAplicar As Decimal
        Get
            Return ObjCreditos_AntDec.ObjValorPro - ObjDebitos_AntDec.ObjValorPro
        End Get
    End Property
    Friend ReadOnly Property DecAnticipoReintegrado As Decimal
        Get
            Dim ldecAntRei = 0D
            For Each lobjNovAnt As ClsNovedadAnticipo In ColNovedadesAnt
                If lobjNovAnt.ObjIdTipoNovedad_NovAntByt.ObjValorPro = EnuTipoNov.EnuDbAntDev Then
                    ldecAntRei += lobjNovAnt.ObjValor_NovAntDec.ObjValorPro
                ElseIf lobjNovAnt.ObjIdTipoNovedad_NovAntByt.ObjValorPro = EnuTipoNov.EnuRDbAntDev Then
                    ldecAntRei -= lobjNovAnt.ObjValor_NovAntDec.ObjValorPro
                End If
            Next
            Return ldecAntRei
        End Get
    End Property
    Friend ReadOnly Property BlnEsAnticipoAjuste As Boolean
        Get
            Dim lblnEsAntAjuste As Boolean
            lblnEsAntAjuste = (ObjIdTipoDocOrigen_AntByt.ObjValorPro = EnuTipoDocOri.EnuNotaAjuste)
            Return lblnEsAntAjuste
        End Get
    End Property
    Friend ReadOnly Property ObjPredioAgrAnt As ClsPredio
        Get
            If IsNothing(MobjPredioAgrAnt) Then
                If Not String.IsNullOrEmpty(ObjIdPredioAgrupador_AntStr.ToString) AndAlso
                        ObjIdPredioAgrupador_AntStr.ToString <> GCSTRSINPA Then
                    MobjPredioAgrAnt = New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
                    Dim lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil,
                            ObjIdPredioAgrupador_AntStr.ObjValorPro}
                    MobjPredioAgrAnt.SAbra(lobjValorLlave)
                    If Not MobjPredioAgrAnt.BlnExiste Then
                        MobjPredioAgrAnt = Nothing
                    End If
                End If
            End If
            Return MobjPredioAgrAnt
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SVacie()
        MyBase.SVacie()
        MdtbNovedadesAnt = Nothing
        McolNovedadesAnt = Nothing
        MdtbNotasCon = Nothing
        McolNotasCon = Nothing
        McolNovedadesRev = Nothing
    End Sub
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        Try
            GobjPanDat.SControleProcesoObj(True)
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                SNumereObj()
            End If
            ClsPanorama.SActualiceCol(ColNovedadesAnt)
            ObjValor_AntDec.SValide()
            ObjFechaCreacionDtm.SValide()
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
    End Sub
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjIdAnticipoEnt.ToString
        End Get
    End Property
#End Region
#Region "Procedimientos del objeto"
    Private Sub SNumereObj()
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            Dim lshrIdAnticipo As Short
            lshrIdAnticipo = ClsPanorama.FobjUltimaIdNumericaObjeto(SstrNombreTabla,
                    ObjIdAnticipoEnt.StrNombreCampoBD, ObjIdAnticipoEnt.EnuTipoValor,
                    ClsOrionCop.StrFiltroUbicacion) + 1
            ObjIdAnticipoEnt.ObjValorPro = lshrIdAnticipo
            For Each lobjNovedadAnt As ClsNovedadAnticipo In McolNovedadesAnt
                lobjNovedadAnt.ObjIdAnticipo_NovEnt.ObjValorPro = lshrIdAnticipo
            Next
        End If
    End Sub
    Friend Sub SReverseAnticipo(aenuTipoDocOrigen As EnuTipoDocOri,
                                astrPrefDocOri As String,
                                aentIdDocOri As Integer,
                                adtmFechaReversion As Date)
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
            HblnEsAnulable = True
            EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
        End If
        ObjAnuladoBln.ObjValorPro = True
        SReverseNotasCon(astrPrefDocOri, aentIdDocOri, adtmFechaReversion)
        SReverseNovedadesAnt(aenuTipoDocOrigen, astrPrefDocOri, aentIdDocOri, adtmFechaReversion)
        SActualice(True)
    End Sub
    Private Sub SReverseNotasCon(astrPrefNotaRRC As String, aentIdNotaRRC As Integer,
                                adtmFechaNotaRRC As Date)
        For Each lobjNotaCon As ClsNotaCon In ColNotasCon
            lobjNotaCon.SReverse(astrPrefNotaRRC, aentIdNotaRRC, adtmFechaNotaRRC)
        Next
    End Sub
    Friend Sub SReverseNovedadesAnt(aenuTipoDocOrigen As EnuTipoDocOri,
                                    astrPrefijoDocOrigen As String,
                                    aentIdDocOrigen As Integer,
                                    adtmfechaDocOrigen As Date)
        For Each lobjNovAnt As ClsNovedadAnticipo In ColNovedadesAnt
            SGenereNovReversionAntRecibido(aenuTipoDocOrigen, lobjNovAnt, astrPrefijoDocOrigen,
                                           aentIdDocOrigen, adtmfechaDocOrigen)
        Next
        If Not IsNothing(McolNovedadesRev) Then
            For Each lobjNovRev As ClsNovedadAnticipo In McolNovedadesRev
                lobjNovRev.ObjIdNovedadAntShr.ObjValorPro = McolNovedadesAnt.Count + 1
                McolNovedadesAnt.Add(lobjNovRev, McolNovedadesAnt.Count + 1)
            Next
            SActualiceValoresAnt()
        End If
    End Sub
    Friend Function FblnAntReintegrado() As Boolean
        Dim lblnAntRev As Boolean
        lblnAntRev = (DecAnticipoReintegrado > 0)
        Return lblnAntRev
    End Function
    Friend Function FblnAnticipoReversado() As Boolean
        Dim lblnReversado = False
        For Each lobjNovAnt As ClsNovedadAnticipo In ColNovedadesAnt
            If lobjNovAnt.ObjIdTipoDocOrigen_NovAntByt.ObjValorPro = EnuTipoDocOri.EnuNotaRevCr Then
                lblnReversado = True
                Exit For
            End If
        Next
        Return lblnReversado
    End Function
    Friend Function FstrAliasCon() As String
        Dim lstrAliasCon = String.Empty
        If ObjPredioAgrAnt IsNot Nothing Then
            lstrAliasCon = ObjPredioAgrAnt.ObjAliasContStr.ToString
        End If
        If String.IsNullOrEmpty(lstrAliasCon) Then
            lstrAliasCon = ObjIdCliente_AntDbl.ToString
        End If
        Return lstrAliasCon
    End Function
#End Region
#Region "Manejo Novedades Anticipo"
    Friend ReadOnly Property ColNovedadesAnt As Collection
        Get
            If IsNothing(McolNovedadesAnt) OrElse McolNovedadesAnt.Count = 0 Then
                McolNovedadesAnt = New Collection
                If ObjIdAnticipoEnt.BlnEsValido AndAlso EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                    SCargueDtbNovedadesAnt()
                    For Each ldrwNovAnt As DataRow In MdtbNovedadesAnt.Rows
                        Dim lobjNovedadAnt As New ClsNovedadAnticipo(Me, ldrwNovAnt)
                        lobjNovedadAnt.SLeaValores(True)
                        McolNovedadesAnt.Add(lobjNovedadAnt, lobjNovedadAnt.ObjIdNovedadAntShr.ToString)
                    Next
                End If
            End If
            Return McolNovedadesAnt
        End Get
    End Property
    Friend ReadOnly Property DtbNovedadesAnt As DataTable
        Get
            SCargueDtbNovedadesAnt()
            SComplementeTablaNov()
            Return MdtbNovedadesAnt
        End Get
    End Property
    Private Sub SCargueDtbNovedadesAnt()
        If IsNothing(MdtbNovedadesAnt) Then
            Dim lstrIdAnticipo = "0"
            If Not IsNothing(ObjIdAnticipoEnt.ObjValorPro) Then
                lstrIdAnticipo = ObjIdAnticipoEnt.ToString
            End If
            Dim lstrIndice = {{StrCampoCarpeta, "ASC"},
                              {StrCampoCentroUtil, "ASC"},
                              {ClsIdAnticipo_NovEnt.SstrNombreCampoBd, "ASC"},
                              {ClsIdNovedadAntShr.SstrNombreCampoBd, "ASC"}}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsIdAnticipo_NovEnt.SstrNombreCampoBd &
                    " = " & lstrIdAnticipo
            Dim lstrCamposSelect() = {"*", "'' AS DocOrigen", "'' AS NroDocOrigen", "'' AS Detalle"}
            MdtbNovedadesAnt = ClsPanorama.FdtbDataTable(ClsNovedadAnticipo.SstrNombreTabla, lstrCamposSelect,
                    lstrIndice, lstrFiltro)
        End If
    End Sub
    Friend Sub SGenereNovedadAntRecibido(adtmFechaNovedadAnt As Date, astrIdCuentaDb As String,
                adecValorNovAnt As Decimal)
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            If IsNothing(McolNovedadesAnt) Then
                McolNovedadesAnt = ColNovedadesAnt
            End If
            Dim lobjNovedadAnt = FobjNuevaNovedadAnt()
            With lobjNovedadAnt
                .ObjFechaNovedadAntDtm.ObjValorPro = adtmFechaNovedadAnt
                .ObjIdCuentaCr_NovAntStr.ObjValorPro = GobjParametros.ObjIdCtaAnticiposRecibidosStr.ObjValorPro
                .ObjIdCuentaDb_NovAntStr.ObjValorPro = astrIdCuentaDb
                .ObjIdDocOrigen_NovAntEnt.ObjValorPro = ObjIdDocOrigen_AntEnt.ObjValorPro
                .ObjIdTipoDocOrigen_NovAntByt.ObjValorPro = ObjIdTipoDocOrigen_AntByt.ObjValorPro
                .ObjPrefijoDocOrigen_NovAntStr.ObjValorPro = ObjPrefijoDocOrigen_AntStr.ObjValorPro
                .ObjIdTipoNovedad_NovAntByt.ObjValorPro = EnuTipoNov.EnuCrAntRec
                .ObjValor_NovAntDec.ObjValorPro = adecValorNovAnt
            End With
            McolNovedadesAnt.Add(lobjNovedadAnt)
            ObjValor_AntDec.ObjValorPro += lobjNovedadAnt.ObjValor_NovAntDec.ObjValorPro
            ObjCreditos_AntDec.ObjValorPro += lobjNovedadAnt.ObjValor_NovAntDec.ObjValorPro
        End If
    End Sub
    Friend Sub SGenereNovedadAntDevuelto(aobjNotaDevAnt As ClsNotaDevAnt)
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
            HblnEsModificable = True
            EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
        End If
        Dim lobjNovedadAnt = FobjNuevaNovedadAnt()
        With lobjNovedadAnt
            .ObjFechaNovedadAntDtm.ObjValorPro = aobjNotaDevAnt.ObjFecha_NotaDevAntDtm.ObjValorPro
            .ObjIdAnticipo_NovEnt.ObjValorPro = ObjIdAnticipoEnt.ObjValorPro
            .ObjIdCuentaCr_NovAntStr.ObjValorPro = GobjParametros.ObjIdCtaCajaStr.ObjValorPro
            .ObjIdCuentaDb_NovAntStr.ObjValorPro = GobjParametros.ObjIdCtaAnticiposRecibidosStr.ObjValorPro
            .ObjIdDocOrigen_NovAntEnt.ObjValorPro = aobjNotaDevAnt.ObjIdNotaDevAntEnt.ObjValorPro
            .ObjIdTipoDocOrigen_NovAntByt.ObjValorPro = EnuTipoDocOri.EnuNotaDevAnt
            .ObjIdTipoNovedad_NovAntByt.ObjValorPro = EnuTipoNov.EnuDbAntDev
            .ObjPrefijoDocOrigen_NovAntStr.ObjValorPro = String.Empty
            .ObjValor_NovAntDec.ObjValorPro = aobjNotaDevAnt.ObjValor_NotaDevAntDec.ObjValorPro
        End With
        ObjDebitos_AntDec.ObjValorPro += lobjNovedadAnt.ObjValor_NovAntDec.ObjValorPro
        ColNovedadesAnt.Add(lobjNovedadAnt)
    End Sub
    Friend Sub SGenereNovedadAntAplicado(aobjNotaCon As ClsNotaCon)
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
            EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
        End If
        Dim ldecValorAntAplicado As Decimal = aobjNotaCon.FdecValorAntAplicado
        ObjDebitos_AntDec.ObjValorPro += ldecValorAntAplicado
        If IsNothing(McolNovedadesAnt) Then
            McolNovedadesAnt = ColNovedadesAnt
        End If
        Dim lobjNovedadAnt = FobjNuevaNovedadAnt()
        With lobjNovedadAnt
            .ObjFechaNovedadAntDtm.ObjValorPro = aobjNotaCon.ObjFecha_NotaConDtm.ObjValorPro
            .ObjIdAnticipo_NovEnt.ObjValorPro = ObjIdAnticipoEnt.ObjValorPro
            .ObjIdTipoNovedad_NovAntByt.ObjValorPro = EnuTipoNov.EnuDbAntApl
            .ObjIdCuentaCr_NovAntStr.ObjValorPro = "*"
            .ObjIdCuentaDb_NovAntStr.ObjValorPro = GobjParametros.ObjIdCtaAnticiposRecibidosStr.ObjValorPro
            .ObjIdDocOrigen_NovAntEnt.ObjValorPro = aobjNotaCon.ObjIdNotaConEnt.ObjValorPro
            .ObjIdTipoDocOrigen_NovAntByt.ObjValorPro = EnuTipoDocOri.EnuNotaCon
            .ObjPrefijoDocOrigen_NovAntStr.ObjValorPro = aobjNotaCon.ObjPrefijo_NotaConStr.ObjValorPro
            .ObjValor_NovAntDec.ObjValorPro = ldecValorAntAplicado
        End With
        McolNovedadesAnt.Add(lobjNovedadAnt)
    End Sub
    Private Sub SGenereNovReversionAntRecibido(aenuTipoDocOrigen As EnuTipoDocOri,
                                               aobjNovAnt As ClsNovedadAnticipo,
                                               astrPrefDocOri As String,
                                               aentIdDocOrigen As Integer,
                                               adtmFecha As Date)
        Dim lobjNovAntRev As ClsNovedadAnticipo = FobjNuevaNovedadAnt()
        Dim lenuTipoNov As EnuTipoNov = aobjNovAnt.ObjIdTipoNovedad_NovAntByt.ObjValorPro
        Dim lenuTipoNovRev As EnuTipoNov = EnuTipoNov.None
        If lenuTipoNov <> EnuTipoNov.EnuDbAntDev AndAlso lenuTipoNov <> EnuTipoNov.EnuRDbAntDev Then
            If IsNothing(McolNovedadesRev) Then
                McolNovedadesRev = New Collection
            End If
            lenuTipoNovRev = ClsOrionCop.FenuTipoNovContraria(lenuTipoNov)
            With lobjNovAntRev
                .ObjIdAnticipo_NovEnt.ObjValorPro = ObjIdAnticipoEnt.ObjValorPro
                .ObjFechaNovedadAntDtm.ObjValorPro = adtmFecha
                .ObjIdCuentaCr_NovAntStr.ObjValorPro = aobjNovAnt.ObjIdCuentaDb_NovAntStr.ObjValorPro
                .ObjIdCuentaDb_NovAntStr.ObjValorPro = aobjNovAnt.ObjIdCuentaCr_NovAntStr.ObjValorPro
                .ObjIdDocOrigen_NovAntEnt.ObjValorPro = aentIdDocOrigen
                .ObjIdTercero_NovAntDbl.ObjValorPro = aobjNovAnt.ObjIdTercero_NovAntDbl.ObjValorPro
                .ObjAliasCont_NovAntStr.ObjValorPro = aobjNovAnt.ObjAliasCont_NovAntStr.ObjValorPro
                .ObjIdTipoDocOrigen_NovAntByt.ObjValorPro = aenuTipoDocOrigen
                .ObjIdTipoNovedad_NovAntByt.ObjValorPro = lenuTipoNovRev
                .ObjIdUsuarioAnuloStr.ObjValorPro = GstrIdUsuario
                .ObjOrigenInstanciaAnuloStr.ObjValorPro = GstrOrigenActual
                .ObjOrigenInstanciaStr.ObjValorPro = GstrOrigenActual
                .ObjPrefijoDocOrigen_NovAntStr.ObjValorPro = astrPrefDocOri
                .ObjValor_NovAntDec.ObjValorPro = aobjNovAnt.ObjValor_NovAntDec.ObjValorPro
                If DecAnticipoReintegrado > 0 Then
                    Throw New ErrorInesperadoPanLException("Anticipo reintegrado")
                End If
            End With
            McolNovedadesRev.Add(lobjNovAntRev)
        End If
    End Sub
    Friend Sub SAnuleAntDevuelto(astrPrefNotaDevAnt As String,
            aentIdNotaDevAnt As Integer)
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
            HblnEsModificable = True
            EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
        End If
        Dim lobjNovAntDev = FobjNovedadDevAnt(astrPrefNotaDevAnt, aentIdNotaDevAnt)
        ' Anula la novedad
        lobjNovAntDev.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
        lobjNovAntDev.ObjAnuladoBln.ObjValorPro = True
        lobjNovAntDev.ObjIdUsuarioAnuloStr.ObjValorPro = GstrIdUsuario
        lobjNovAntDev.ObjOrigenInstanciaAnuloStr.ObjValorPro = GstrOrigenActual
        ObjDebitos_AntDec.ObjValorPro -= lobjNovAntDev.ObjValor_NovAntDec.ObjValorPro
        lobjNovAntDev.ObjValor_NovAntDec.ObjValorPro = 0
    End Sub
    ' Procedimiento ejecutado solo en la actualizaciòn a la versiòn 09.02.237
    Friend Sub SGenereNovReversionAntRec(aenuTipoDocorigen As EnuTipoDocOri,
                                         astrPrefDocOri As String,
                                         aentIdDocorigen As Integer,
                                         adtmFechaDocOri As Date)
        Dim ldecVlr = 0D
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
            HblnEsAnulable = True
            EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
        End If
        ObjAnuladoBln.ObjValorPro = True
        For Each lobjNovAnt As ClsNovedadAnticipo In ColNovedadesAnt
            If lobjNovAnt.ObjIdTipoNovedad_NovAntByt.ObjValorPro = EnuTipoNov.EnuCrAntRec Then
                ldecVlr = lobjNovAnt.ObjValor_NovAntDec.ObjValorPro
                SGenereNovReversionAntRecibido(aenuTipoDocorigen, lobjNovAnt, astrPrefDocOri,
                                               aentIdDocorigen, adtmFechaDocOri)
                Exit For
            End If
        Next
        If Not IsNothing(McolNovedadesRev) Then
            For Each lobjNovRev As ClsNovedadAnticipo In McolNovedadesRev
                lobjNovRev.ObjIdNovedadAntShr.ObjValorPro = McolNovedadesAnt.Count + 1
                McolNovedadesAnt.Add(lobjNovRev, McolNovedadesAnt.Count + 1)
            Next
        End If
        ObjValor_AntDec.ObjValorPro = ldecVlr
        ObjCreditos_AntDec.ObjValorPro = ldecVlr
        ObjDebitos_AntDec.ObjValorPro = ldecVlr
        ObjFechaCreacionDtm.SValide()
        If FblnEstanTodosOk() Then
            SActualice(True)
        End If
    End Sub
    Private Sub SActualiceValoresAnt()
        Dim lenuTipoNov As EnuTipoNov
        Dim ldecVlrNov As Decimal
        If McolNovedadesRev.Count > 0 Then
            For Each lobjNov As ClsNovedadAnticipo In McolNovedadesRev
                lenuTipoNov = lobjNov.ObjIdTipoNovedad_NovAntByt.ObjValorPro
                ldecVlrNov = lobjNov.ObjValor_NovAntDec.ObjValorPro
                If lenuTipoNov <> EnuTipoNov.EnuDbAntDev Then
                    Select Case lenuTipoNov
                        Case EnuTipoNov.EnuRCrAntRec
                            ObjDebitos_AntDec.ObjValorPro += ldecVlrNov
                        Case EnuTipoNov.EnuRDbAntApl
                            ObjCreditos_AntDec.ObjValorPro += ldecVlrNov
                    End Select
                End If
            Next
        End If
    End Sub
    Friend Function FobjNuevaNovedadAnt() As ClsNovedadAnticipo
        Dim lobjNovedadAnt As ClsNovedadAnticipo = Nothing
        Try
            SCargueDtbNovedadesAnt()
            Dim ldrwNewNovedadAnt = MdtbNovedadesAnt.NewRow
            Dim lblnModificoPermisos = False
            lobjNovedadAnt = New ClsNovedadAnticipo(Me, ldrwNewNovedadAnt)
            With lobjNovedadAnt
                If Not CType(.EnuPermisosObj And EnuPermisosDef.enuCrear, Boolean) Then
                    .EnuPermisosObj += EnuPermisosDef.enuCrear
                    lblnModificoPermisos = True
                End If
                .SCreeObj(Nothing)
                .ObjAnuladoBln.ObjValorPro = False
                .ObjFechaCreacionDtm.ObjValorPro = Date.Now
                .ObjIdCarpeta_NovAntShr.ObjValorPro = GshrIdCarpeta
                .ObjIdCentroUtil_NovAntShr.ObjValorPro = GshrIdCentroUtil
                .ObjAliasCont_NovAntStr.ObjValorPro = FstrAliasCon()
                .ObjIdTercero_NovAntDbl.ObjValorPro = ObjIdCliente_AntDbl.ObjValorPro
                .ObjIdCuentaCr_NovAntStr.ObjValorPro = GobjParametros.ObjIdCtaAnticiposRecibidosStr.ObjValorPro
                If lblnModificoPermisos Then
                    .EnuPermisosObj -= EnuPermisosDef.enuCrear
                End If
            End With
        Catch ex As ErrorInesperadoPanDatException
            Throw
        Catch ex As ErrorInesperadoPanLException
            Throw
        Catch ex As PanDatException
            Throw
        Catch ex As PanLException
            Throw
        Catch ex As ArgumentNullException
            Throw
        Catch ex As Exception
            Throw
        End Try
        Return lobjNovedadAnt
    End Function
    Private Function FobjNovedadDevAnt(astrPrefNotaDevAnt As String,
            aentIdNotaDevAnt As Integer) As ClsNovedadAnticipo
        Dim lobjNovAntDev As ClsNovedadAnticipo = Nothing
        For Each lobjNovAnt As ClsNovedadAnticipo In ColNovedadesAnt
            If lobjNovAnt.ObjIdTipoNovedad_NovAntByt.ObjValorPro = EnuTipoNov.EnuDbAntDev Then
                If lobjNovAnt.ObjIdTipoDocOrigen_NovAntByt.ObjValorPro =
                        EnuTipoDocOri.EnuNotaDevAnt AndAlso
                        lobjNovAnt.ObjPrefijoDocOrigen_NovAntStr.ObjValorPro = astrPrefNotaDevAnt AndAlso
                        lobjNovAnt.ObjIdDocOrigen_NovAntEnt.ObjValorPro = aentIdNotaDevAnt Then
                    lobjNovAntDev = lobjNovAnt
                    Exit For
                End If
            End If
        Next
        Return lobjNovAntDev
    End Function
    Private Sub SComplementeTablaNov()
        Dim lstrPrefDocOri As String, lentIdDocOri As Integer, lstrNroDocOri As String
        Dim lenuTipoDocOri As EnuTipoDocOri
        Dim lenuTipoNovAnt As EnuTipoNov
        Dim lstrDetalle = String.Empty, lstrDocOrigen As String
        Dim ldrwNovedades = MdtbNovedadesAnt.Select
        For Each ldrwNovedad As DataRow In ldrwNovedades
            lenuTipoDocOri = ClsPanorama.FobjValorCampo(ldrwNovedad(ClsIdTipoDocOrigen_AntByt.SstrNombreCampoBd),
                    EnuTipoValor.enuByte)
            lenuTipoNovAnt = ClsPanorama.FobjValorCampo(ldrwNovedad(ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd),
                    EnuTipoValor.enuByte)
            lstrPrefDocOri = ClsPanorama.FobjValorCampo(ldrwNovedad(ClsPrefijoDocOrigen_NovAntStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            lentIdDocOri = ClsPanorama.FobjValorCampo(ldrwNovedad(ClsIdDocOrigen_NovAntEnt.SstrNombreCampoBd),
                    EnuTipoValor.enuInteger)
            lstrNroDocOri = ClsPanorama.FstrNumeroDcto(lstrPrefDocOri, lentIdDocOri)
            ldrwNovedad("NroDocOrigen") = lstrNroDocOri
            lstrDocOrigen = ClsOrionCop.FstrDocOrigenNovedad(lenuTipoDocOri)
            Select Case lenuTipoNovAnt
                Case EnuTipoNov.EnuCrAntRec
                    lstrDetalle = "Anticipo recibido"
                Case EnuTipoNov.EnuDbAntApl
                    lstrDetalle = "Anticipo aplicado"
                Case EnuTipoNov.EnuDbAntDev
                    lstrDetalle = "Anticipo reintegrado"
                Case EnuTipoNov.EnuRCrAntRec
                    lstrDetalle = "Anticipo recibido reversado"
                Case EnuTipoNov.EnuRDbAntApl
                    lstrDetalle = "Anticipo aplicado reversado"
                Case EnuTipoNov.EnuRDbAntDev
                    lstrDetalle = "Anticipo reintegrado reversado"
            End Select
            ldrwNovedad("DocOrigen") = lstrDocOrigen
            ldrwNovedad("NroDocOrigen") = lstrNroDocOri
            ldrwNovedad("Detalle") = lstrDetalle
        Next
    End Sub
#End Region
#Region "Notas Contables"
    ''' <summary>
    ''' Devuelve una colección de las notas contables mediante las cuales se aplicó
    ''' el presente anticipo
    ''' </summary>
    Friend ReadOnly Property ColNotasCon As Collection
        Get
            Dim lstrPrefNco As String, lentIdNco As Integer
            If IsNothing(McolNotasCon) OrElse McolNotasCon.Count = 0 Then
                McolNotasCon = New Collection
                If ObjIdAnticipoEnt.BlnEsValido AndAlso EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                    SCargueDtbNotasCon()
                    For Each ldrwNovNot As DataRow In MdtbNotasCon.Rows
                        Dim lobjNotaCon As New ClsNotaCon()
                        lstrPrefNco = ClsPanorama.FobjValorCampo(ldrwNovNot(
                                ClsPrefijo_NotaConStr.SstrNombreCampoBd), EnuTipoValor.enuString)
                        lentIdNco = ClsPanorama.FobjValorCampo(ldrwNovNot(
                                ClsIdNotaConEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
                        lobjNotaCon.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefNco, lentIdNco})
                        McolNotasCon.Add(lobjNotaCon, lobjNotaCon.ObjIdNotaConEnt.ToString)
                    Next
                End If
            End If
            Return McolNotasCon
        End Get
    End Property
    Private Sub SCargueDtbNotasCon()
        If IsNothing(MdtbNotasCon) Then
            Dim lstrIndice = {{ClsPrefijo_NotaConStr.SstrNombreCampoBd, "ASC"},
                              {ClsIdNotaConEnt.SstrNombreCampoBd, "ASC"}}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsIdAnticipo_NotaConEnt.SstrNombreCampoBd &
                    " = " & ObjIdAnticipoEnt.ToString
            Dim lstrCamposSelect() = {ClsPrefijo_NotaConStr.SstrNombreCampoBd, ClsIdNotaConEnt.SstrNombreCampoBd}
            MdtbNotasCon = ClsPanorama.FdtbDataTable(ClsNotaCon.SstrNombreTabla, lstrCamposSelect,
                    lstrIndice, lstrFiltro)
        End If
    End Sub
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsCreditos_AntDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Creditos"
    Private ReadOnly MobjPadre As ClsAnticipo = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Creditos Anticipo"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Decimal.MaxValue,
                BlnEsRequerido, HenuTipoValor)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
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
Friend Class ClsDebitos_AntDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Debitos"
    Private ReadOnly MobjPadre As ClsAnticipo = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "DebitosAnticipo"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = 0
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, Decimal.MaxValue,
                    BlnEsRequerido, EnuTipoValor.enuDecimal)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
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
Friend Class ClsFechaAnticipoDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaAnticipo"
    Private ReadOnly MobjPadre As ClsAnticipo = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "FechaAnticipo"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MyBase.SVaciePropiedad()
        HobjValorNew = GCDTMFECHANULA
        HobjValorPro = GCDTMFECHANULA
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = True
        If Not GblnActualizandoApp Then
            Dim ldtmFechaMin = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
            Dim ldtmFechaMax = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax, BlnEsRequerido)
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
        If Not IsNothing(HobjValorPro) Then
            Return Format(HobjValorPro, GCSTRFMTFECHASIMPLE)
        Else
            Return Format(GCDTMFECHANULA, GCSTRFMTFECHASIMPLE)
        End If
    End Function
End Class
Friend Class ClsIdAnticipoEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdAnticipo"
    Private ReadOnly MobjPadre As ClsAnticipo = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdAnticipo"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 2
    End Sub
    Public Overrides Sub SValide()
        HstrMens = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue,
                BlnEsRequerido, EnuTipoValor)
        If HblnEsValido Then
            If Not BlnLeyendoOrigen Then
                If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                    Dim lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
                    If MobjPadre.FblnExisteLlave(lobjValorLlave) Then
                        HstrMens = "Un Anticipo con el número de identificación ingresado, '" &
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
Friend Class ClsIdCliente_AntDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTerceroCliente"
    Private ReadOnly MobjPadre As ClsAnticipo = Nothing
    Private MobjCliente As ClsCliente = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdTerceroCliente_Anticipo"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCDBLMINTERC, GCDBLMAXTERC, BlnEsRequerido)
        If HblnEsValido Then
            MobjCliente = MobjPadre.ObjPadre
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = (HobjValorNew = MobjCliente.ObjIdClienteDbl.ObjValorPro)
            Else
                Dim lobjLlavePrincipal() = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
                If IsNothing(MobjCliente) Then
                    MobjCliente = New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
                    MobjCliente.SAbra(lobjLlavePrincipal)
                    If Not MobjCliente.BlnExiste Then
                        Throw New ErrorInesperadoPanLException("Anticipo sin Cliente valido!")
                    End If
                End If
            End If
        End If
        If Not HblnEsValido Then
            MobjCliente = Nothing
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
Friend Class ClsIdDocOrigen_AntEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdDocOrigen"
    Private ReadOnly MobjPadre As ClsAnticipo = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Id Doc Origen Anticipo"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue, BlnEsRequerido,
                EnuTipoValor.enuInteger)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                If Not GblnActualizandoApp Then
                    HblnEsValido = (HobjValorNew = HobjValorOriginal)
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
Friend Class ClsIdPredioAgrupador_AntStr
    'Herencia
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdPredioAgrupador"
    Private ReadOnly MobjPadre As ClsAnticipo = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdPredioAgrupador_Anticipo"
        HenuTipoValor = EnuTipoValor.enuString
        HshrLongitud = 20
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = False
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud,
                BlnEsRequerido)
        If HblnEsValido AndAlso HobjValorNew <> "" Then
            Dim lobjLlavePrincipal() As Object = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
                lobjPredio.SAbra(lobjLlavePrincipal)
                HblnEsValido = lobjPredio.BlnExiste
                If HblnEsValido Then
                    HblnEsValido = (lobjPredio.ObjIdPredioStr.ObjValorPro = lobjPredio.ObjIdPredioAgrupadorStr.ObjValorPro)
                End If
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
        If Not IsNothing(ObjValorPro) Then
            Return HobjValorPro.ToString()
        Else
            Return ""
        End If
    End Function
End Class
Friend Class ClsIdPredio_AntStr
    'Herencia
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdPredio"
    Private ReadOnly MobjPadre As ClsAnticipo = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdPredio_Ant"
        HenuTipoValor = EnuTipoValor.enuString
        HshrLongitud = 20
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = False
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = String.Empty
        HblnEsRequerido = (MobjPadre.ObjIdTipoDocOrigen_AntByt.ObjValorPro =
                           EnuTipoDocOri.EnuNotaAjuste)
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                If HobjValorNew <> "" Then
                    Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
                    Dim lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
                    lobjPredio.SAbra(lobjValorLlave)
                    HblnEsValido = lobjPredio.BlnExiste
                End If
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
        If Not IsNothing(ObjValorPro) Then
            Return HobjValorPro.ToString
        Else
            Return ""
        End If
    End Function
End Class
Friend Class ClsIdTipoDocOrigen_AntByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTipoDocOrigen"
    Private ReadOnly MobjPadre As ClsAnticipo = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdTipoDocOrigen"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuTipoDocOri.EnuReciboCaja,
                EnuTipoDocOri.EnuNotaAjuste, BlnEsRequerido)
        If HblnEsValido Then
            Dim lobjPadre As ClsAnticipo = ObjPadre
            If lobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            Else
                HblnEsValido = (HobjValorNew = EnuTipoDocOri.EnuReciboCaja OrElse
                                HobjValorNew = EnuTipoDocOri.EnuNotaAjuste)
            End If
        End If
    End Sub
    Private Sub EPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            MobjPadre.ObjPrefijoDocOrigen_AntStr.SValide()
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
Friend Class ClsPrefijoDocOrigen_AntStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PrefijoDocOrigen"
    Private ReadOnly MobjPadre As ClsAnticipo = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Prefijo Doc Origen Anticipo"
        HshrLongitud = 5
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then
            HobjValorNew = String.Empty
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud, BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                If MobjPadre.ObjIdTipoDocOrigen_AntByt.ObjValorPro = EnuTipoDocOri.EnuNotaAjuste Then
                    HblnEsValido = (HobjValorNew = String.Empty)
                Else
                    HblnEsValido = (HobjValorNew = GobjParametros.FstrPrefijoDoc(
                            EnuTipoDocOri.EnuReciboCaja))
                End If
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
Friend Class ClsServicios_AntStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "ServiciosAnt"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Servicio Anticipo"
        HshrLongitud = 50
        HblnEsRequerido = True
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud, BlnEsRequerido)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If Not IsNothing(HobjValorPro) Then
            Return HobjValorPro
        Else
            Return ""
        End If
    End Function
End Class
Friend Class ClsValor_AntDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "ValorAnticipo"
    Private ReadOnly MobjPadre As ClsAnticipo = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "ValorAnticipo"
        HenuTipoValor = EnuTipoValor.enuDecimal
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0.01, Decimal.MaxValue,
                BlnEsRequerido, HenuTipoValor)
        If Not HblnEsValido Then
            If HobjValorNew = 0 Then
                HblnEsValido = MobjPadre.ObjAnuladoBln.ObjValorPro
            End If
        Else
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                If Not GblnActualizandoApp Then
                    HblnEsValido = (HobjValorNew = HobjValorOriginal)
                End If
            End If
        End If
    End Sub
    Private Sub ClsValor_AntDec_evnPosSetValor(sender As Object, e As ClsPanEventArgs) Handles Me.EvnPosSetValor
        MobjPadre.ObjCreditos_AntDec.SValide()
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
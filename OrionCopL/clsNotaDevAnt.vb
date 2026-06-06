Friend Class ClsNotaDevAnt
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriNotasDevAnt"
    '
    Private MobjPredioAgrNDevAnt As ClsPredio = Nothing
    Private MobjClienteNota As clsCliente = Nothing
    Private MdtbNovedades As DataTable = Nothing
#End Region
#Region "Constructores"
    ''' <summary>
    ''' Instancia un objeto Nota Crédito en modo único
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
    ''' Instancia un objeto Nota Crédito en modo navegable
    ''' </summary>
    Public Sub New(astrPref As String)
        HobjPadre = Nothing
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion
        lstrFiltro &= " AND " & ClsPrefijo_NotaDevAntStr.SstrNombreCampoBd & " = '" & astrPref & "'"
        HcolFiltros.Add(lstrFiltro)
        Dim lstrCamposSelect = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsPrefijo_NotaDevAntStr.SstrNombreCampoBd, ClsIdNotaDevAntEnt.SstrNombreCampoBd}
        HblnEsSuprimible = False
        HblnEsModificable = False
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
        henuTipoObjeto = enuModoInstanciaObjDef.enuDeColeccion
        hblnEsAnulable = False
        HblnEsSuprimible = False
        hblnEsModificable = False
        '
        drwRegistroActual = adrwObjeto
        DtbTablaColeccion = drwRegistroActual.Table
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
            Return EnuIdClasesPanDef.enuNotaDevAnt
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Nota Devolución Anticipos"
        End Get
    End Property
    Friend Overrides ReadOnly Property HstrNombreObj As String
        Get
            Return Chr(34) & "Nro. " & StrNumeroNotaDevAnt & Chr(34)
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjComentario_NotaDevAntStr As New ClsComentario_NotaDevAntStr(Me)
    Friend ReadOnly Property ObjFechaAnulacion_NotaDevAntDtm As New ClsFechaAnulacion_NotaDevAntDtm(Me)
    Friend ReadOnly Property ObjFecha_NotaDevAntDtm As New ClsFecha_NotaDevAntDtm(Me)
    Friend ReadOnly Property ObjIdAnticipo_NotaDevAntEnt As New ClsIdAnticipo_NotaDevAntEnt(Me)
    Friend ReadOnly Property ObjIdCarpeta_NotaDevAntShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_NotaDevAntShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdCliente_NotaDevAntDbl As New ClsIdCliente_NotaDevAntDbl(Me)
    Friend ReadOnly Property ObjIdNotaDevAntEnt As New ClsIdNotaDevAntEnt(Me)
    Friend ReadOnly Property ObjIdPredioAgrupador_NotaDevAntStr As New ClsIdPredioAgrupador_NotaDevAntStr(Me)
    Friend ReadOnly Property ObjIdUsuario_NotaDevAntStr As New ClsIdUsuarioStr(Me)
    Friend ReadOnly Property ObjPrefijo_NotaDevAntStr As New ClsPrefijo_NotaDevAntStr(Me)
    Friend ReadOnly Property ObjValor_NotaDevAntDec As New ClsValor_NotaDevAntDec(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjAnuladoBln)
                HcolPropiedades.Add(ObjFechaCreacionDtm)
                HcolPropiedades.Add(ObjIdUsuarioAnuloStr)
                HcolPropiedades.Add(ObjOrigenInstanciaStr)
                HcolPropiedades.Add(ObjComentario_NotaDevAntStr)
                HcolPropiedades.Add(ObjFechaAnulacion_NotaDevAntDtm)
                HcolPropiedades.Add(ObjFecha_NotaDevAntDtm)
                HcolPropiedades.Add(ObjIdAnticipo_NotaDevAntEnt)
                HcolPropiedades.Add(ObjIdCarpeta_NotaDevAntShr)
                HcolPropiedades.Add(ObjIdCentroUtil_NotaDevAntShr)
                HcolPropiedades.Add(ObjIdCliente_NotaDevAntDbl)
                HcolPropiedades.Add(ObjIdNotaDevAntEnt)
                HcolPropiedades.Add(ObjIdPredioAgrupador_NotaDevAntStr)
                HcolPropiedades.Add(ObjIdUsuario_NotaDevAntStr)
                HcolPropiedades.Add(ObjPrefijo_NotaDevAntStr)
                HcolPropiedades.Add(ObjValor_NotaDevAntDec)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    ''' <summary>
    ''' Devuelve un string compuesto por el prefijo de la nota y el id de la nota separados por un
    ''' guion. Si no existe el prefijo devuelve solo el id de la nota
    ''' </summary>
    ''' <value></value>
    Friend ReadOnly Property StrNumeroNotaDevAnt As String
        Get
            Dim lstrNumeroNotaDevAnt As String = ClsPanorama.FstrNumeroDcto(ObjPrefijo_NotaDevAntStr.ObjValorPro,
                    ObjIdNotaDevAntEnt.ObjValorPro)
            Return lstrNumeroNotaDevAnt
        End Get
    End Property
    Friend Property ObjClienteNota As clsCliente
        Get
            Dim lobjValorLlave As Object() = {ObjIdCarpeta_NotaDevAntShr.ObjValorPro,
                ObjIdCentroUtil_NotaDevAntShr.ObjValorPro, ObjIdCliente_NotaDevAntDbl.ObjValorPro}
            MobjClienteNota = New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
            MobjClienteNota.SAbra(lobjValorLlave)
            Return MobjClienteNota
        End Get
        Set(value As ClsCliente)
            MobjClienteNota = value
            HobjPadre = MobjClienteNota
        End Set
    End Property
    Friend ReadOnly Property ObjPredioAgrNDevAnt As ClsPredio
        Get
            If IsNothing(MobjPredioAgrNDevAnt) Then
                If Not String.IsNullOrEmpty(ObjIdPredioAgrupador_NotaDevAntStr.ToString) AndAlso
                        ObjIdPredioAgrupador_NotaDevAntStr.ToString <> GCSTRSINPA Then
                    MobjPredioAgrNDevAnt = New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
                    Dim lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil,
                            ObjIdPredioAgrupador_NotaDevAntStr.ObjValorPro}
                    MobjPredioAgrNDevAnt.SAbra(lobjValorLlave)
                    If Not MobjPredioAgrNDevAnt.BlnExiste Then
                        MobjPredioAgrNDevAnt = Nothing
                    End If
                End If
            End If
            Return MobjPredioAgrNDevAnt
        End Get
    End Property
#End Region
#Region "Anticipo"
    Friend ReadOnly Property ObjAnticipo As ClsAnticipo
        Get
            Return ObjIdAnticipo_NotaDevAntEnt.ObjAnticipo
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Public Overrides Function FblnEsAnulable() As Boolean
        Dim lblnEsAnulable = BlnEsAnulable
        If lblnEsAnulable Then
            If GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro Then
                lblnEsAnulable = (Date.Today <= GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo)
            End If
        End If
        If lblnEsAnulable Then
            Dim lstrPeriodoNota = ClsPanorama.FstrPeriodo(ObjFecha_NotaDevAntDtm.ObjValorPro)
            Dim lstrPeriodoActual = GobjParametros.ObjAnoActual.StrIdPeriodoActual
            lblnEsAnulable = (lstrPeriodoNota = lstrPeriodoActual)
        End If
        Return lblnEsAnulable
    End Function
    Protected Overrides Function SAnuleEnObj() As Boolean
        Dim lblnAnulado = FblnEsAnulable()
        If lblnAnulado Then
            ObjAnuladoBln.ObjValorPro = True
            ObjIdUsuarioAnuloStr.ObjValorPro = GstrIdUsuario
            ObjOrigenInstanciaAnuloStr.ObjValorPro = GstrOrigenActual
            If GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro Then
                ObjFechaAnulacion_NotaDevAntDtm.ObjValorPro = Now
            Else
                If GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo < Date.Today Then
                    ObjFechaAnulacion_NotaDevAntDtm.ObjValorPro = Now.AddDays(-Date.Today.Day)
                Else
                    ObjFechaAnulacion_NotaDevAntDtm.ObjValorPro = Now
                End If
            End If
        End If
        ' Anticipo
        If Not IsNothing(ObjAnticipo) AndAlso ObjAnticipo.BlnExiste Then
            ObjAnticipo.SAnuleAntDevuelto(ObjPrefijo_NotaDevAntStr.ObjValorPro,
                ObjIdNotaDevAntEnt.ObjValorPro)
            ObjAnticipo.SActualice(True)
        Else
            Throw New ErrorInesperadoPanLException("Anticipo no existe!")
        End If
        ObjValor_NotaDevAntDec.ObjValorPro = 0
        Return lblnAnulado
    End Function
    Protected Overrides Sub SInicialiceObj()
        ObjAnuladoBln.ObjValorPro = False
        ObjFechaAnulacion_NotaDevAntDtm.ObjValorPro = GCDTMFECHANULA
        ObjFechaCreacionDtm.ObjValorPro = Date.Now
        ObjIdUsuarioAnuloStr.ObjValorPro = String.Empty
        ObjOrigenInstanciaStr.ObjValorPro = GstrOrigenActual
        ObjIdCarpeta_NotaDevAntShr.ObjValorPro = GshrIdCarpeta
        ObjIdCentroUtil_NotaDevAntShr.ObjValorPro = GshrIdCentroUtil
        ObjIdUsuario_NotaDevAntStr.ObjValorPro = GstrIdUsuario
        ObjFechaCreacionDtm.ObjValorPro = Date.Now
        ObjIdUsuarioAnuloStr.ObjValorPro = String.Empty
        ObjOrigenInstanciaAnuloStr.ObjValorPro = String.Empty
    End Sub
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        Dim lblnNoHayError = False
        GobjPanDat.SControleProcesoObj(True)
        Try
            GobjPanDat.SInicialiceTransaccion()
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                SNumereObj()
                ObjAnticipo.SGenereNovedadAntDevuelto(Me)
                ObjAnticipo.SActualice(True)
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
        MdtbNovedades = Nothing
        MobjClienteNota = Nothing
    End Sub
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ObjIdNotaDevAntEnt.ToString
        End Get
    End Property
#End Region
#Region "Procedimientos del objeto"
    Private Sub SNumereObj()
        If enuEstadoActualizacion = enuEstadoObjetoDef.enuCreando Then
            Dim lentIdNotaDevAnt As Integer
            lentIdNotaDevAnt = clsPanorama.fobjUltimaIdNumericaObjeto(sstrNombreTabla,
                    objIdNotaDevAntEnt.strNombreCampoBD, objIdNotaDevAntEnt.enuTipoValor,
                    clsOrionCop.strFiltroUbicacion) + 1
            ObjPrefijo_NotaDevAntStr.ObjValorPro =
                    GobjParametros.FstrPrefijoDoc(EnuTipoDocOri.EnuNotaDevAnt)
            ObjIdNotaDevAntEnt.ObjValorPro = lentIdNotaDevAnt
        End If
    End Sub
    Friend Function FstrAliasCon() As String
        Dim lstrAliasCon = String.Empty
        If ObjPredioAgrNDevAnt IsNot Nothing Then
            lstrAliasCon = ObjPredioAgrNDevAnt.ObjAliasContStr.ToString
        End If
        If String.IsNullOrEmpty(lstrAliasCon) Then
            lstrAliasCon = ObjIdCliente_NotaDevAntDbl.ToString
        End If
        Return lstrAliasCon
    End Function
#End Region
#Region "Novedades"
    Friend ReadOnly Property DtbNovedades As DataTable
        Get
            sCargueDtbNovedades()
            sComplementeTablaNov()
            Return mdtbNovedades
        End Get
    End Property
    Private Sub SCargueDtbNovedades()
        If IsNothing(MdtbNovedades) Then
            Dim lstrIdNotaDevAnt = ObjIdNotaDevAntEnt.ToString
            If String.IsNullOrEmpty(lstrIdNotaDevAnt) Then lstrIdNotaDevAnt = "0"
            Dim lstrIndice = {{ClsIdNovedadAntShr.SstrNombreCampoBd, "ASC"}}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsIdTipoDocOrigenByt.SstrNombreCampoBd &
                    " = " & EnuTipoDocOri.EnuNotaDevAnt & " AND " &
                    ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd & " = '" & ObjPrefijo_NotaDevAntStr.ObjValorPro &
                    "' AND " & ClsIdDocOrigen_NovAntEnt.SstrNombreCampoBd & " = " & lstrIdNotaDevAnt
            Dim lstrCamposSelect() = {"*", "'' AS ConceptoNovedad"}
            MdtbNovedades = ClsPanorama.FdtbDataTable(ClsNovedadAnticipo.SstrNombreTabla, lstrCamposSelect, lstrIndice,
                    lstrFiltro)
        End If
    End Sub
    Private Sub SComplementeTablaNov()
        Dim lstrConceptoNovedad = String.Empty
        For Each ldrwNovedad As DataRow In MdtbNovedades.Rows
            Dim lenuTipoNovedad As EnuTipoNov =
                    ClsPanorama.FobjValorCampo(ldrwNovedad(ClsIdTipoNovedad_NovAntByt.SstrNombreCampoBd),
                    EnuTipoValor.enuByte)
            If lenuTipoNovedad = EnuTipoNov.EnuDbAntDev Then
                lstrConceptoNovedad = "Anticipo reintegrado"
            ElseIf lenuTipoNovedad = EnuTipoNov.EnuRDbAntDev Then
                lstrConceptoNovedad = "Reversado Anticipo reintegrado"
            End If
            ldrwNovedad("ConceptoNovedad") = lstrConceptoNovedad
        Next
    End Sub
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsComentario_NotaDevAntStr
    Inherits clsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Comentario"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Comentario"
        HshrLongitud = 500
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 3, ShrLongitud, BlnEsRequerido)
        If Not HblnEsValido AndAlso (IsNothing(HobjValorNew) OrElse TypeOf HobjValorNew Is String) Then
            HstrMens = "El Detalle es obligatorio y debe tener mas de tres letras!!"
            SNotifiqueDatInv()
        End If
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
Friend Class ClsFechaAnulacion_NotaDevAntDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaAnulacion"
    Private ReadOnly MobjPadre As ClsNotaDevAnt = ObjPadre
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "FechaAnulacion_NotaDevAnt"
        HenuTipoValor = EnuTipoValor.enuDateTime
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = False
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MyBase.SVaciePropiedad()
        HobjValorNew = GCDTMFECHANULA
        HobjValorPro = HobjValorNew
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = Not IsNothing(HobjValorNew) AndAlso IsDate(HobjValorNew)
        If HobjValorNew <> GCDTMFECHANULA Then
            If HblnEsValido Then
                Dim ldtmFechaMin As Date = GCDTMFECHANULA
                Dim ldtmFechaMax As Date = GCDTMFECHANULA
                If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                    HblnEsValido = MobjPadre.FblnEsAnulable
                    If HblnEsValido Then
                        HblnEsValido = MobjPadre.ObjAnuladoBln.ObjValorPro
                        If HblnEsValido Then
                            If Today > GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo Then
                                ldtmFechaMin = Today.AddDays(-Today.Day)
                                ldtmFechaMax = Now.AddDays(-Now.Day)
                            Else
                                ldtmFechaMin = Date.Today
                                ldtmFechaMax = Now
                            End If
                        End If
                        HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax,
                                BlnEsRequerido)
                    End If
                ElseIf MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                    HblnEsValido = (HobjValorNew = HobjValorOriginal)
                Else
                    HblnEsValido = False
                End If
            End If
        End If
    End Sub
    Private Sub EPosSetValor() Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando AndAlso BlnEsValido Then
            If HobjValorPro <> GCDTMFECHANULA Then
                MobjPadre.ObjValor_NotaDevAntDec.SValide()
            End If
        End If
    End Sub
    Public Overrides Function ToString() As String
        If Not IsNothing(HobjValorPro) Then
            Return Format(HobjValorPro, GCSTRFMTFECHASIMPLE)
        Else
            Return Format(GCDTMFECHANULA, GCSTRFMTFECHASIMPLE)
        End If
    End Function
End Class
Friend Class ClsFecha_NotaDevAntDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaNota"
    Private ReadOnly MobjPadre As ClsNotaDevAnt = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "FechaNotaDevAnt"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MyBase.SVaciePropiedad()
        HobjValorNew = Date.Today
        HobjValorPro = Date.Today
    End Sub
    Public Overrides Sub SValide()
        Dim lobjValorIng = HobjValorNew
        Dim ldtmFechaMin = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
        Dim ldtmFechaMax = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
        If ldtmFechaMax > Date.Today Then
            ldtmFechaMax = Now
        End If
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            HstrMens = String.Empty
            HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax, BlnEsRequerido)
            If Not HblnEsValido Then
                HstrMens = "La fecha de la Devolución está por fuera del Período Actual!"
            ElseIf Not IsDate(HobjValorNew) Then
                HstrMens = "La Fecha ingresada, '" &
                        lobjValorIng.ToString & "',  no es una Fecha valida!"
            End If
            If Not String.IsNullOrEmpty(HstrMens) Then
                SNotifiqueDatInv()
            End If
        Else
            HblnEsValido = (HobjValorNew = HobjValorOriginal)
        End If
    End Sub
    Friend Function FblnFechaAntEsValida() As Boolean
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando AndAlso HblnEsValido Then
            If MobjPadre.ObjIdAnticipo_NotaDevAntEnt.ObjAnticipo.BlnExiste Then
                Dim ldtmFechaAnt As Date = MobjPadre.ObjIdAnticipo_NotaDevAntEnt.ObjAnticipo.ObjFechaAnticipoDtm.ObjValorPro
                HblnEsValido = (HobjValorNew >= ldtmFechaAnt)
            End If
        End If
        Return HblnEsValido
    End Function
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
Friend Class ClsIdAnticipo_NotaDevAntEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdAnticipo"
    Private ReadOnly MobjPadre As ClsNotaDevAnt = Nothing
    Private MobjAnticipo As New ClsAnticipo(EnuModoInstanciaObjDef.enuUnico)
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdAnticipo_NotaDevAnt"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue, BlnEsRequerido,
                EnuTipoValor.enuInteger)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            Else
                MobjAnticipo = ObjAnticipo
                HblnEsValido = MobjAnticipo.BlnExiste AndAlso MobjPadre.ObjIdPredioAgrupador_NotaDevAntStr.ObjValorPro =
                        MobjAnticipo.ObjIdPredioAgrupador_AntStr.ObjValorPro
                If HblnEsValido Then
                    HblnEsValido = MobjPadre.ObjFecha_NotaDevAntDtm.FblnFechaAntEsValida()
                    If Not HblnEsValido Then
                        HstrMens = "La fecha de la Devolución es anterior a la fecha del Anticipo!"
                        SNotifiqueDatInv()
                    End If
                End If
            End If
        End If
    End Sub
    Friend ReadOnly Property ObjAnticipo() As ClsAnticipo
        Get
            If Not MobjAnticipo.BlnExiste OrElse MobjAnticipo.ObjIdAnticipoEnt.ObjValorPro <> HobjValorNew Then
                Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
                MobjAnticipo.SAbra(lobjValorLlave)
            End If
            Return MobjAnticipo
        End Get
    End Property
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
        MobjPadre.ObjFecha_NotaDevAntDtm.SValide()
        MobjPadre.ObjIdPredioAgrupador_NotaDevAntStr.SValide()
    End Sub
End Class
Friend Class ClsIdCliente_NotaDevAntDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTerceroCliente"
    Private ReadOnly MobjPadre As ClsNotaDevAnt = Nothing
    Private MobjCliente As ClsCliente = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdTerceroCliente_NotaDevAnt"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        Dim lobjValorIng = HobjValorNew
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCDBLMINTERC, GCDBLMAXTERC, BlnEsRequerido)
        HstrMens = String.Empty
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                If IsNothing(MobjCliente) Then
                    MobjCliente = New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
                End If
                Dim lobjLlavePrincipal() = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
                MobjCliente.SAbra(lobjLlavePrincipal)
                MobjPadre.ObjClienteNota = MobjCliente
                If Not MobjCliente.BlnExiste Then
                    HblnEsValido = False
                    MobjCliente.SVacie()
                    HstrMens = "La Id. del Cliente ingresada, '" & lobjValorIng.ToString &
                            "',  no es valida!"
                End If
            End If
        Else
            HstrMens = "La Id. del Cliente ingresada, '" & lobjValorIng.ToString & "',  no es válida!"
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
Friend Class ClsIdNotaDevAntEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdNotaDevAnt"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdNotaDevAnticipo"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
        HblnEsLlave = True
        HbytPosicionLlave = 3
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 1, Integer.MaxValue, BlnEsRequerido,
                        EnuTipoValor)
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
Friend Class ClsIdPredioAgrupador_NotaDevAntStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdPredioAgrupador"
    Private ReadOnly MobjPadre As ClsNotaDevAnt = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdPredioAgrupador NotaDevAnt"
        HenuTipoValor = EnuTipoValor.enuString
        HshrLongitud = 20
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud,
                BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = (HobjValorNew <> My.Resources.Ninguno)
                If HblnEsValido Then
                    HblnEsValido = String.IsNullOrEmpty(HobjValorNew) OrElse
                            ClsOrionCop.FblnExistePredio(HobjValorNew)
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
Friend Class ClsPrefijo_NotaDevAntStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PrefijoNota"
    Private ReadOnly MobjPadre As ClsCBObjetoPan = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "PrefijoNotaCon"
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
                HblnEsValido = (HobjValorNew =
                        GobjParametros.FstrPrefijoDoc(EnuTipoDocOri.EnuNotaDevAnt))
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
Friend Class ClsValor_NotaDevAntDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Valor"
    Private ReadOnly MobjPadre As ClsNotaDevAnt = Nothing
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
        HstrMens = String.Empty
        If Not HblnEsValido Then
            If HobjValorNew = 0 Then
                HblnEsValido = MobjPadre.ObjAnuladoBln.ObjValorPro
            End If
            If Not HblnEsValido Then
                HstrMens = "El Valor de la Nota debe ser mayor a Cero!"
            End If
        Else
            HblnEsValido = HobjValorNew - Int(HobjValorNew) = 0
            If HblnEsValido Then
                If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                    HblnEsValido = (HobjValorNew = HobjValorOriginal)
                Else
                    HblnEsValido = (HobjValorNew <= MobjPadre.ObjIdAnticipo_NotaDevAntEnt.ObjAnticipo.DecAnticipoPorAplicar)
                    If Not HblnEsValido Then
                        HstrMens = "El valor de la Devolución es mayor al Saldo del Anticipo!"
                    End If
                End If
            Else
                HstrMens = "El Valor ingresado debe ser sin Centavos!"
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
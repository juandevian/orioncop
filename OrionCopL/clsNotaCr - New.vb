Imports System.Text
Imports System.Drawing
Imports System.Windows.Media.Imaging
Imports ThoughtWorks.QRCode.Codec
Friend Class ClsNotaCr
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriNotasCr"
    ' Variables de modulo
    Private MobjClienteNotaCr As ClsCliente = Nothing
    Private MobjPredioAgrNCr As ClsPredio = Nothing
    Private McolItemsNotaCr As Collection = Nothing
    Private MdtbNovedades As DataTable = Nothing
    Private McolNovedades As Collection = Nothing
    Private MobjNotaReversionRC As ClsNotaReversionCr = Nothing
    Private MstrNroNotaRCr As String = String.Empty
    Private MdecDsctoPorValor As Decimal = 0
    Private MDecValorCrMaxCap As Decimal = -1
    Private MdecValorCrMaxInt As Decimal = -1
    Private MecValorCrMaxIva As Decimal = -1
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
        HenuTipoObjeto = EnuModoInstanciaObjDef.EnuUnico
        HcolTablas.Add(MCSTRNOMBRETABLA)
        HcolCamposSelect.Add({"*"})
    End Sub
    ''' <summary>
    ''' Instancia un objeto Nota Crédito en modo navegable
    ''' </summary>
    Public Sub New(astrPref As String)
        HobjPadre = Nothing
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion
        lstrFiltro &= " AND " & ClsPrefijo_NotaCrStr.SstrNombreCampoBd & " = '" & astrPref & "'"
        HcolFiltros.Add(lstrFiltro)
        Dim lstrCamposSelect = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsPrefijo_NotaCrStr.SstrNombreCampoBd, ClsIdNotaCrEnt.SstrNombreCampoBd}
        HblnEsSuprimible = False
        HblnEsModificable = False
        HcolTablas.Add(MCSTRNOMBRETABLA)
        HcolCamposSelect.Add(lstrCamposSelect)
    End Sub
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Administrador al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwNotaCr">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Public Sub New(aobjPadre As ClsCliente, adrwNotaCr As DataRow)
        HobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.EnuDeColeccion
        HblnEsAnulable = False
        HblnEsSuprimible = False
        HblnEsModificable = False
        HblnEsCreable = False
        '
        DrwRegistroActual = adrwNotaCr
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
            Return EnuIdClasesPanDef.EnuNotaCr
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Nota Crédito"
        End Get
    End Property
    Friend Overrides ReadOnly Property HstrNombreObj As String
        Get
            Return Chr(34) & "Nro. " & StrNumeroNotaCr & Chr(34)
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjComentario_NotaCrStr As New ClsComentario_NotaCrStr(Me)
    Friend ReadOnly Property ObjCUDEStr As New ClsCUDEStr(Me)
    Friend ReadOnly Property ObjCUDocStr As New ClsCUDocStr(Me)
    Friend ReadOnly Property ObjFecha_NotaCrDtm As New ClsFecha_NotaCrDtm(Me)
    Friend ReadOnly Property ObjFechaAnulacion_NotaCrDtm As New ClsFechaAnulacion_NotaCrDtm(Me)
    Friend ReadOnly Property ObjIdCarpeta_NotaCrShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_NotaCrShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdEstadoEDocEnt As New ClsIdEstadoEDocEnt(Me)
    Friend ReadOnly Property ObjIdNotaCrEnt As New ClsIdNotaCrEnt(Me)
    Friend ReadOnly Property ObjIdPredioAgrupador_NotaCrStr As New ClsIdPredioAgrupador_NotaCrStr(Me)
    Friend ReadOnly Property ObjIdCliente_NotaCrDbl As New ClsIdCliente_NotaCrDbl(Me)
    Friend ReadOnly Property ObjIdTipoNotaCrByt As New ClsIdTipoNotaCrByt(Me)
    Friend ReadOnly Property ObjIdUsuario_NotaCrStr As New ClsIdUsuarioStr(Me)
    Friend ReadOnly Property ObjModoNotaCrByt As New ClsModoNotaCrByt(Me)
    Friend ReadOnly Property ObjPrefijo_NotaCrStr As New ClsPrefijo_NotaCrStr(Me)
    Friend ReadOnly Property ObjValor_NotaCrDec As New ClsValor_NotaCrDec(Me)
    Friend ReadOnly Property ObjVerEFacEnt As New ClsVerEFacEnt(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjAnuladoBln)
                HcolPropiedades.Add(ObjFechaCreacionDtm)
                HcolPropiedades.Add(ObjIdUsuarioAnuloStr)
                HcolPropiedades.Add(ObjOrigenInstanciaStr)
                HcolPropiedades.Add(ObjOrigenInstanciaAnuloStr)
                HcolPropiedades.Add(ObjComentario_NotaCrStr)
                HcolPropiedades.Add(ObjCUDEStr)
                HcolPropiedades.Add(ObjCUDocStr)
                HcolPropiedades.Add(ObjFecha_NotaCrDtm)
                HcolPropiedades.Add(ObjFechaAnulacion_NotaCrDtm)
                HcolPropiedades.Add(ObjIdCarpeta_NotaCrShr)
                HcolPropiedades.Add(ObjIdCentroUtil_NotaCrShr)
                HcolPropiedades.Add(ObjIdCliente_NotaCrDbl)
                HcolPropiedades.Add(ObjIdEstadoEDocEnt)
                HcolPropiedades.Add(ObjIdNotaCrEnt)
                HcolPropiedades.Add(ObjIdPredioAgrupador_NotaCrStr)
                HcolPropiedades.Add(ObjIdTipoNotaCrByt)
                HcolPropiedades.Add(ObjIdUsuario_NotaCrStr)
                HcolPropiedades.Add(ObjModoNotaCrByt)
                HcolPropiedades.Add(ObjPrefijo_NotaCrStr)
                HcolPropiedades.Add(ObjValor_NotaCrDec)
                HcolPropiedades.Add(ObjVerEFacEnt)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    Friend Property BlnAnulandoFac As Boolean = False
    Friend Property BlnGenerandoRec As Boolean = False
    ''' <summary>
    ''' Devuelve uns string compuesto por el prefijo de la nota y el id de la nota separados por un
    ''' guion. Si no existe el prefijo devuelve solo el id de la factura
    ''' </summary>
    ''' <value></value>
    Friend ReadOnly Property StrNumeroNotaCr As String
        Get
            Dim lstrNumeroNotaCr As String = ClsPanorama.FstrNumeroDcto(
                    ObjPrefijo_NotaCrStr.ObjValorPro, ObjIdNotaCrEnt.ObjValorPro)
            Return lstrNumeroNotaCr
        End Get
    End Property
    Friend Property ObjClienteNotaCr As ClsCliente
        Get
            Dim lobjValorLlave As Object() = {ObjIdCarpeta_NotaCrShr.ObjValorPro,
                ObjIdCentroUtil_NotaCrShr.ObjValorPro, ObjIdCliente_NotaCrDbl.ObjValorPro}
            MobjClienteNotaCr = New ClsCliente(EnuModoInstanciaObjDef.EnuUnico)
            MobjClienteNotaCr.SAbra(lobjValorLlave)
            Return MobjClienteNotaCr
        End Get
        Set(value As ClsCliente)
            MobjClienteNotaCr = value
            HobjPadre = MobjClienteNotaCr
        End Set
    End Property
    Friend ReadOnly Property ObjPredioAgrNCr As ClsPredio
        Get
            If IsNothing(MobjPredioAgrNCr) Then
                If Not String.IsNullOrEmpty(ObjIdPredioAgrupador_NotaCrStr.ToString()) AndAlso
                        ObjIdPredioAgrupador_NotaCrStr.ToString() <> String.Empty Then
                    MobjPredioAgrNCr = New ClsPredio(EnuModoInstanciaObjDef.EnuUnico)
                    Dim lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil,
                            ObjIdPredioAgrupador_NotaCrStr.ObjValorPro}
                    MobjPredioAgrNCr.SAbra(lobjValorLlave)
                    If Not MobjPredioAgrNCr.BlnExiste Then
                        MobjPredioAgrNCr = Nothing
                    End If
                End If
            End If
            Return MobjPredioAgrNCr
        End Get
    End Property
    ''' <summary>
    ''' Devuelve la factura que afecta el primer item de la nota. Para facturación elecrrónica solo puede
    ''' existir una npta por factura.
    ''' </summary>
    ''' <returns></returns>
    Friend ReadOnly Property ObjFacturaAfectada As ClsFactura
        Get
            Dim lobjItemNCr As ClsItemNotaCr = ColItemsNotaCr(1)
            Return lobjItemNCr.ObjFactura
        End Get
    End Property
    Friend ReadOnly Property ColNovedades As Collection
        Get
            If IsNothing(McolNovedades) Then
                McolNovedades = New Collection
                SCargueDtbNovedades()
                If Not IsNothing(MdtbNovedades) AndAlso MdtbNovedades.Rows.Count > 0 Then
                    For Each ldrwNovedad As DataRow In MdtbNovedades.Rows
                        Dim lobjNovedad As New ClsNovedad(Me, ldrwNovedad)
                        lobjNovedad.SLeaValores(True)
                        McolNovedades.Add(lobjNovedad)
                    Next
                End If
            End If
            Return McolNovedades
        End Get
    End Property
    Friend ReadOnly Property DecDsctosYRetenCausados As Decimal
        Get
            Dim ldecTot = 0D
            Dim lenuTipoNov As EnuTipoNov
            For Each lobjNov As ClsNovedad In ColNovedades
                lenuTipoNov = lobjNov.ObjIdTipoNovedadByt.ObjValorPro
                If lenuTipoNov >= EnuTipoNov.EnuCrDctoCap AndAlso
                        lenuTipoNov <= EnuTipoNov.EnuCrRetCre Then
                    ldecTot += lobjNov.ObjValor_NovDec.ObjValorPro
                End If
            Next
            Return ldecTot
        End Get
    End Property
    Friend ReadOnly Property DecDsctosNota As Decimal
        Get
            Dim ldecTot = 0D
            Dim lenuTipoNov As EnuTipoNov
            For Each lobjNov As ClsNovedad In ColNovedades
                lenuTipoNov = lobjNov.ObjIdTipoNovedadByt.ObjValorPro
                If lenuTipoNov = EnuTipoNov.EnuCrDctoCap OrElse
                        lenuTipoNov = EnuTipoNov.EnuCrDctoInt Then
                    ldecTot += lobjNov.ObjValor_NovDec.ObjValorPro
                End If
            Next
            Return ldecTot
        End Get
    End Property
    Friend ReadOnly Property BlnTieneDsctos As Boolean
        Get
            Dim lblnTiene = False
            For Each lobjItemNotaCr As ClsItemNotaCr In ColItemsNotaCr
                lblnTiene = lobjItemNotaCr.FblnEsDscto
                If lblnTiene Then Exit For
            Next
            Return lblnTiene
        End Get
    End Property
    Friend ReadOnly Property BlnEsRetencion As Boolean
        Get
            Dim lblnEsRet = False
            For Each lobjItemNotaCr As ClsItemNotaCr In ColItemsNotaCr
                lblnEsRet = lobjItemNotaCr.FblnEsRetencion
                If Not lblnEsRet Then Exit For
            Next
            Return lblnEsRet
        End Get
    End Property
    Friend ReadOnly Property StrTipoConceptoDian As String
        Get
            Dim lstrTipoDscto As String
            If ObjIdTipoNotaCrByt.ObjValorPro = EnuTipoNotaCrDef.EnuAnulaFac Then
                lstrTipoDscto = CType(EnuConceptoNotaCrDian.EnuAnulacion, Integer).ToString
            ElseIf FblnCancelaFac() OrElse FblnEsDesctoTotal() Then
                lstrTipoDscto = CType(EnuConceptoNotaCrDian.EnuRebajaDscto, Integer).ToString
            Else
                lstrTipoDscto = CType(EnuConceptoNotaCrDian.EnuRebajaDscto, Integer).ToString
            End If
            Return lstrTipoDscto
        End Get
    End Property
#End Region
#End Region
#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SVacie()
        MyBase.SVacie()
        McolItemsNotaCr = Nothing
        MdtbNovedades = Nothing
        McolNovedades = Nothing
        MobjClienteNotaCr = Nothing
        MobjPredioAgrNCr = Nothing
        MobjNotaReversionRC = Nothing
        MstrNroNotaRCr = String.Empty
        EnuTipoDsctoPorValor = EnuTipoDescuentoDef.None
        MdecDsctoPorValor = 0
        MDecValorCrMaxCap = -1
        MdecValorCrMaxInt = -1
        MecValorCrMaxIva = -1
    End Sub
    Protected Overrides Sub SInicialiceObj()
        ObjAnuladoBln.ObjValorPro = False
        ObjFechaAnulacion_NotaCrDtm.ObjValorPro = GCDTMFECHANULA
        ObjIdCarpeta_NotaCrShr.ObjValorPro = GshrIdCarpeta
        ObjIdCentroUtil_NotaCrShr.ObjValorPro = GshrIdCentroUtil
        ObjIdUsuario_NotaCrStr.ObjValorPro = GstrIdUsuario
        ObjOrigenInstanciaAnuloStr.ObjValorPro = String.Empty
        ObjOrigenInstanciaStr.ObjValorPro = GstrOrigenActual
        ObjIdUsuarioAnuloStr.ObjValorPro = String.Empty
        Dim lstrPrefijo = GobjParametros.FstrPrefijoDoc(EnuTipoDocOri.EnuNotaCr)
        If IsNothing(lstrPrefijo) Then lstrPrefijo = String.Empty
        ObjPrefijo_NotaCrStr.ObjValorPro = lstrPrefijo
        ObjCUDocStr.ObjValorPro = String.Empty
        ObjCUDEStr.ObjValorPro = String.Empty
        ObjVerEFacEnt.ObjValorPro = EnuVerEFac.EnuNinguna
        ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuNoEDoc
    End Sub
    Public Overrides Function FblnEsAnulable() As Boolean
        Dim lblnEsAnulable = BlnEsAnulable
        If lblnEsAnulable Then
            If GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro Then
                lblnEsAnulable = (Date.Today <= GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo)
            End If
        End If
        If lblnEsAnulable Then
            lblnEsAnulable = ObjIdTipoNotaCrByt.ObjValorPro <> EnuTipoNotaCrDef.EnuAnulaFac
        End If
        If lblnEsAnulable Then
            Dim lstrPeriodoNCr = ClsPanorama.FstrPeriodo(ObjFecha_NotaCrDtm.ObjValorPro)
            Dim lstrPeriodoActual = GobjParametros.ObjAnoActual.StrIdPeriodoActual
            lblnEsAnulable = (lstrPeriodoNCr = lstrPeriodoActual)
        End If
        If lblnEsAnulable Then
            If GobjParametros.BlnEFacAutorizado Then
                If Not BlnAnulandoFac Then
                    lblnEsAnulable = Not ObjIdTipoNotaCrByt.ObjValorPro =
                            EnuTipoNotaCrDef.EnuRetenciones
                End If
            End If
        End If
        If lblnEsAnulable Then
            lblnEsAnulable = FblnFechaDocEsPeriodoActual(ObjFecha_NotaCrDtm.ObjValorPro)
        End If
        Return lblnEsAnulable
    End Function
    Private Function FblnEsAnulableNCr() As Boolean
        Dim lblnEsAnulable = BlnEsAnulable
        If lblnEsAnulable Then
            If GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro Then
                lblnEsAnulable = (Date.Today <= GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo)
            End If
        End If
        If lblnEsAnulable Then
            If GobjParametros.BlnEFacAutorizado Then
                If Not BlnAnulandoFac Then
                    lblnEsAnulable = Not (ObjIdTipoNotaCrByt.ObjValorPro =
                            EnuTipoNotaCrDef.EnuRetenciones)
                End If
            End If
        End If
        Return lblnEsAnulable
    End Function
    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        Dim lblnNoHayError = False
        GobjPanDat.SControleProcesoObj(True)
        Try
            GobjPanDat.SInicialiceTransaccion()
            If ColItemsNotaCr.Count = 0 Then
                Throw New ErrorInesperadoPanLException("Inentado actualizar Nota Db sin Items")
            End If
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
                SNumereObj()
                SCompleteItems()
                STipoEsRetencion()
                SComplementeEFac()
                ObjFechaCreacionDtm.ObjValorPro = Date.Now
                ClsPanorama.SActualiceCol(McolItemsNotaCr)
                If Not (BlnAnulandoFac OrElse BlnGenerandoRec) Then
                    SActualiceFacturas()
                End If
                MyBase.SActualice(ablnExigeRequeridos)
            Else
                ClsPanorama.SActualiceCol(McolItemsNotaCr)
                MyBase.SActualice(ablnExigeRequeridos)
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
                Dim lstrMens = "La Nota " & StrNumeroNotaCr & " fue actualizada exitosamente!"
                SLevanteEventoNot(lstrMens, "", 0, EnuSeveridadNot.EnuInformacion)
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    Protected Overrides Function SAnuleEnObj() As Boolean
        Dim lblnAnulado = FblnEsAnulable()
        If lblnAnulado Then
            ObjAnuladoBln.ObjValorPro = True
            ObjIdUsuarioAnuloStr.ObjValorPro = GstrIdUsuario
            ObjOrigenInstanciaAnuloStr.ObjValorPro = GstrOrigenActual
            If GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro Then
                ObjFechaAnulacion_NotaCrDtm.ObjValorPro = Now
            Else
                If GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo < Date.Today Then
                    ObjFechaAnulacion_NotaCrDtm.ObjValorPro = Now.AddDays(-Date.Today.Day)
                Else
                    If GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo < Date.Today Then
                        ObjFechaAnulacion_NotaCrDtm.ObjValorPro = Now.AddDays(-Date.Today.Day)
                    Else
                        ObjFechaAnulacion_NotaCrDtm.ObjValorPro = Now
                    End If
                End If
            End If
            If Not ObjComentario_NotaCrStr.BlnEsValido Then
                ObjComentario_NotaCrStr.ObjValorPro = "Nota Crédito Anulada."
            End If
            If lblnAnulado Then
                ' Nota Reversion Crédito
                SGenereNotaRevCr()
            End If
        End If
        Return lblnAnulado
    End Function
    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return Me.StrNumeroNotaCr
        End Get
    End Property
#End Region
#Region "Procedimientos del objeto"
    Private Sub SNumereObj()
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
            Dim lentIdNotaCr As Integer
            Dim lstrPrefijo = GobjParametros.FstrPrefijoDoc(EnuTipoDocOri.EnuNotaCr)
            If IsNothing(lstrPrefijo) Then lstrPrefijo = String.Empty
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsPrefijo_NotaCrStr.SstrNombreCampoBd & " = '" & lstrPrefijo & "'"
            lentIdNotaCr = ClsPanorama.FobjUltimaIdNumericaObjeto(SstrNombreTabla,
                    ClsIdNotaCrEnt.SstrNombreCampoBd, ObjIdNotaCrEnt.EnuTipoValor,
                    lstrFiltro)
            If lentIdNotaCr < GobjParametros.FentNumeracionInicialDoc(EnuTipoDocOri.EnuNotaCr) Then
                lentIdNotaCr = GobjParametros.FentNumeracionInicialDoc(EnuTipoDocOri.EnuNotaCr)
            End If
            lentIdNotaCr += 1
            ObjPrefijo_NotaCrStr.ObjValorPro = lstrPrefijo
            ObjIdNotaCrEnt.ObjValorPro = lentIdNotaCr
        End If
    End Sub
    Private Sub SCompleteItems()
        Dim lobjItemNotaCr As ClsItemNotaCr = Nothing
        If Not IsNothing(ColItemsNotaCr) Then
            For i As Integer = 1 To ColItemsNotaCr.Count
                lobjItemNotaCr = ColItemsNotaCr(i)
                With lobjItemNotaCr
                    .ObjIdItemNotaCrShr.ObjValorPro = i
                    .ObjPrefijo_ItemNotaCrStr.ObjValorPro = ObjPrefijo_NotaCrStr.ObjValorPro
                    .ObjIdNotaCr_ItemNotaCrEnt.ObjValorPro = ObjIdNotaCrEnt.ObjValorPro
                End With
            Next
        End If
    End Sub
    Private Sub SComplementeEFac()
        If GobjParametros.ObjAutorizaEFacBln.ObjValorPro Then
            If BlnAfectaFrasRegEFac() Then
                ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuNoReg
                ObjVerEFacEnt.ObjValorPro = EnuVerEFac.EnuV2
            End If
        End If
    End Sub
    Private Sub STipoEsRetencion()
        Dim lblnEsRet = ObjIdTipoNotaCrByt.ObjValorPro = EnuTipoNotaCrDef.EnuRetenciones
        If Not lblnEsRet Then
            For Each lobjItemNCr As ClsItemNotaCr In ColItemsNotaCr
                lblnEsRet = lobjItemNCr.BlnEsRetencion
                If Not lblnEsRet Then Exit For
            Next
        End If
        If lblnEsRet Then
            ObjIdTipoNotaCrByt.ObjValorPro = EnuTipoNotaCrDef.EnuRetenciones
            ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuNoEDoc
        End If
    End Sub
    Private Sub SActualiceFacturas()
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
            Dim lstrPrefFact As String, lentIdFact As Integer
            Dim lobjFactura As New ClsFactura()
            Dim lobjValorLlave As Object()
            For Each lobjItemNotaCr As ClsItemNotaCr In McolItemsNotaCr
                lstrPrefFact = lobjItemNotaCr.ObjPrefijoFact_ItemNotaCrStr.ObjValorPro
                lentIdFact = lobjItemNotaCr.ObjIdFactura_ItemNotaCrEnt.ObjValorPro
                lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrPrefFact, lentIdFact}
                lobjFactura.SAbra(lobjValorLlave)
                If Not IsNothing(lobjFactura) Then
                    If lobjFactura.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
                        lobjFactura.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
                    End If
                    lobjFactura.SApliqueNotaCr(lobjItemNotaCr)
                    lobjFactura.SActualice(True)
                End If
            Next
        End If
    End Sub
    Friend Function FblnEsValidaFecha(adtmFechaNotaCr As Date) As Boolean
        Dim lblnEsValida = True
        McolItemsNotaCr = ColItemsNotaCr
        For Each lobjItemNotaCr As ClsItemNotaCr In McolItemsNotaCr
            lblnEsValida = lobjItemNotaCr.FblnEsValidaFechaNotaCr(adtmFechaNotaCr)
            If Not lblnEsValida Then Exit For
        Next
        Return lblnEsValida
    End Function
    Friend Function FdecInteresesMoraPorCausar() As Decimal
        Dim ldecIntMoraPorCausar = 0D
        If Not IsNothing(ObjClienteNotaCr) AndAlso ObjClienteNotaCr.BlnExiste Then
            If ObjIdPredioAgrupador_NotaCrStr.BlnEsValido AndAlso
                    ObjFecha_NotaCrDtm.BlnEsValido Then
                Dim lstrIdPrediosAgr As String() = {ObjIdPredioAgrupador_NotaCrStr.ToString()}
                ldecIntMoraPorCausar = ObjClienteNotaCr.FdecIntMoraPorCausar(lstrIdPrediosAgr,
                        ObjFecha_NotaCrDtm.ObjValorPro)
            End If
        End If
        Return ldecIntMoraPorCausar
    End Function
    ''' <summary>
    ''' Causa intereses de mora al cliente y al predio agrupador del recibo.
    ''' </summary>
    ''' <returns>Devuelve el valor de los intereses de mora causados.</returns>
    ''' <remarks></remarks>
    Friend Function SCauseMoraCliente(ByRef astrMens As String) As Decimal
        Dim ldecIntMoraCausados = 0D
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
            Try
                If Not IsNothing(ObjClienteNotaCr) AndAlso
                        ObjIdPredioAgrupador_NotaCrStr.BlnEsValido AndAlso
                        ObjFecha_NotaCrDtm.BlnEsValido Then
                    Dim lstrIdPrediosAgr As String() = {ObjIdPredioAgrupador_NotaCrStr.ToString()}
                    ldecIntMoraCausados = MobjClienteNotaCr.SCauseMora(lstrIdPrediosAgr,
                                    ObjFecha_NotaCrDtm.ObjValorPro, astrMens)
                End If
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
            End Try
        End If
        Return ldecIntMoraCausados
    End Function
    Friend Function FdecIvaReversado() As Decimal
        Dim ldecIvaRev = 0D
        If ObjIdTipoNotaCrByt.ObjValorPro = EnuTipoNotaCrDef.EnuAnulaFac Then
            For Each lobjitemNCr As ClsItemNotaCr In ColItemsNotaCr
                If lobjitemNCr.ObjEsReversionIvaBln.ObjValorPro Then
                    ldecIvaRev += lobjitemNCr.ObjValor_ItemNotaCrDec.ObjValorPro
                End If
            Next
        End If
        Return ldecIvaRev
    End Function
    ''' <summary>
    ''' Devuelve el valor de la base grabable cuando la nota es una anulación de factura
    ''' </summary>
    ''' <param name="aobjFacturaAnulada">La factura anulada</param>
    ''' <returns></returns>
    Friend Function FdecBaseGrabable(aobjFacturaAnulada As ClsFactura) As Decimal
        Dim ldecBaseGravable = 0D
        If ObjIdTipoNotaCrByt.ObjValorPro = EnuTipoNotaCrDef.EnuAnulaFac Then
            For Each lobjItemFac As ClsItemFactura In aobjFacturaAnulada.ColItemsFactura
                ldecBaseGravable += lobjItemFac.FdecBaseIva
            Next
        End If
        Return ldecBaseGravable
    End Function
    ''' <summary>
    ''' Devuelve el valor descontado que afecta la deuda de capital
    ''' </summary>
    ''' <param name="aenuTipoDscto"></param>
    ''' <returns></returns>
    Friend Function FdecValorDsctoCap() As Decimal
        Dim lenuTipoDscto As EnuTipoDescuentoDef, ldecDscto = 0D
        For Each lobjItemNCr As ClsItemNotaCr In ColItemsNotaCr
            lenuTipoDscto = lobjItemNCr.ObjIdTipoDscto_ItemNotaCrByt.ObjValorNuevo
            If lenuTipoDscto = EnuTipoDescuentoDef.EnuDsctoCapital OrElse
                    (lenuTipoDscto >= EnuTipoDescuentoDef.EnuReteFuente AndAlso
                    lenuTipoDscto <= EnuTipoDescuentoDef.EnuDsctoPP) Then
                ldecDscto += lobjItemNCr.ObjValor_ItemNotaCrDec.ObjValorPro
            End If
        Next
        Return ldecDscto
    End Function
    Private Function FblnCancelaFac() As Boolean
        Dim lblnCancela = False
        For Each lobjItemNCr As ClsItemNotaCr In ColItemsNotaCr
            If lobjItemNCr.ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro =
                    EnuTipoDescuentoDef.EnuCancelaIva Then
                lblnCancela = True
                Exit For
            End If
        Next
        Return lblnCancela
    End Function
    Friend Function FarlCorreosNCr() As ArrayList
        Dim larlListaCorreos As New ArrayList
        Dim lstrCorreoCli As String, lstrCorreoPredio As String
        If ObjClienteNotaCr.ObjRecibeDocsPorEmailBln.ObjValorPro Then
            lstrCorreoCli = ObjClienteNotaCr.ObjEmailStr.ToString
            If Not String.IsNullOrEmpty(lstrCorreoCli) Then
                larlListaCorreos.Add(lstrCorreoCli)
            End If
        End If
        If Not IsNothing(ObjPredioAgrNCr) Then
            lstrCorreoPredio = ObjPredioAgrNCr.ObjEmailAdiStr.ToString
            If Not String.IsNullOrEmpty(lstrCorreoPredio) Then
                larlListaCorreos.Add(lstrCorreoPredio)
            End If
        End If
        Return larlListaCorreos
    End Function
    Friend Function FstrAliasCon() As String
        Dim lstrAliasCon = String.Empty
        If ObjPredioAgrNCr IsNot Nothing Then
            lstrAliasCon = ObjPredioAgrNCr.ObjAliasContStr.ToString
        End If
        If String.IsNullOrEmpty(lstrAliasCon) Then
            lstrAliasCon = ObjIdCliente_NotaCrDbl.ToString()
        End If
        Return lstrAliasCon
    End Function
    Friend Function FbytQRNcr() As Byte()
        Dim lentColorFondoQR As Integer = Color.FromArgb(255, 255, 255, 255).ToArgb()
        Dim lentColorQR As Integer = Color.FromArgb(255, 0, 0, 0).ToArgb()
        Dim lqreQRNdb As New QRCodeEncoder With {
            .QRCodeEncodeMode = QRCodeEncoder.ENCODE_MODE.BYTE,
            .QRCodeScale = Int32.Parse(4),
            .QRCodeErrorCorrect = QRCodeEncoder.ERROR_CORRECTION.H,
            .QRCodeVersion = 0,
            .QRCodeBackgroundColor = System.Drawing.Color.FromArgb(lentColorFondoQR),
            .QRCodeForegroundColor = System.Drawing.Color.FromArgb(lentColorQR)
       }
        Dim lstrNit As String = GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual.
                ObjIdTerceroCentroUtilDbl.ToString
        Dim lbytQR As Byte() = Array.Empty(Of Byte)()
        Dim ldtmFecNCr As Date = ObjFecha_NotaCrDtm.ObjValorPro
        Dim lstrFecNCr As String = Year(ldtmFecNCr).ToString &
                Format(Month(ldtmFecNCr), "00") & Format(Day(ldtmFecNCr), "00") &
                Format(Hour(ldtmFecNCr), "00") & Format(Minute(ldtmFecNCr), "00") &
                Format(Second(ldtmFecNCr), "00")
        Dim lstrNdb As String = "Número:" & StrNumeroNotaCr & vbCrLf &
                "Fecha:" & lstrFecNCr & vbCrLf &
                "Nit:" & lstrNit & vbCrLf &
                "DocAdq:" & ObjIdCliente_NotaCrDbl.ToString & vbCrLf &
                "ValNcr:" & Format(ObjValor_NotaCrDec.ObjValorPro, "#0.00") & vbCrLf &
                "ValIva:" & Format(0, "#0.00") & vbCrLf &
                "ValOtroIm:" & "0.00" & vbCrLf &
                "ValNcrIm:" & Format(ObjValor_NotaCrDec.ObjValorPro, "#0.00") & vbCrLf
        If GobjParametros.ObjIdProveedorEFacEnt.ObjValorPro > EnuProveedorEFac.None Then
            lstrNdb &= "CUDE:" & ObjCUDEStr.ObjValorPro
        End If
        Try
            Dim lbmiQRNdb As New BitmapImage
            lbmiQRNdb.BeginInit()
            Dim lbtmQR As Bitmap = lqreQRNdb.Encode(lstrNdb, System.Text.Encoding.UTF8)
            lbytQR = FBytQR(lbtmQR)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly)
        End Try
        Return lbytQR
    End Function
    Friend Function FblnDeudaSuspendida() As Boolean
        Dim lblnDeudaSuspendida As Boolean
        If ObjPredioAgrNCr IsNot Nothing Then
            lblnDeudaSuspendida = ObjPredioAgrNCr.ObjIdEstadoDeuda_PredioByt.ObjValorPro >=
                    EnuEstadoDeudaDef.EnuPerdida
        Else
            lblnDeudaSuspendida = ObjClienteNotaCr.ObjIdEstadoDeudaByt.ObjValorPro >=
                    EnuEstadoDeudaDef.EnuPerdida
        End If
        Return lblnDeudaSuspendida
    End Function
#End Region
#Region "Manejo descuentos por valor"
    Friend Property EnuTipoDsctoPorValor As EnuTipoDescuentoDef = EnuTipoDescuentoDef.None
    Friend Property DecDsctoPorValor As Decimal
        Get
            Return MdecDsctoPorValor
        End Get
        Set(value As Decimal)
            Dim lblnEsValido = ClsPanorama.FblnEsValidoNumero(value, 1.0, Decimal.MaxValue,
                    True, EnuTipoValor.EnuDecimal)
            If lblnEsValido Then
                MdecDsctoPorValor = value
            Else
                MdecDsctoPorValor = 0
            End If
        End Set
    End Property
    Friend ReadOnly Property DecValorCrMaxCap As Decimal
        Get
            If MDecValorCrMaxCap = -1 Then
                MDecValorCrMaxCap = FdecValorCrMaxCap(MdecValorCrMaxInt, MecValorCrMaxIva)
            End If
            Return MDecValorCrMaxCap
        End Get
    End Property
    Friend ReadOnly Property DecValorCrMaxInt As Decimal
        Get
            If MdecValorCrMaxInt = -1 Then
                MDecValorCrMaxCap = FdecValorCrMaxCap(MdecValorCrMaxInt, MecValorCrMaxIva)
            End If
            Return MdecValorCrMaxInt
        End Get
    End Property
    Friend ReadOnly Property DecValorCrMaxIva As Decimal
        Get
            If MecValorCrMaxIva = -1 Then
                MDecValorCrMaxCap = FdecValorCrMaxCap(MdecValorCrMaxInt, MecValorCrMaxIva)
            End If
            Return MecValorCrMaxIva
        End Get
    End Property
    Private Function FdecValorCrMaxCap(ByRef adecValorCrMaxInt As Decimal,
            ByRef adecValorCrMaxIva As Decimal) As Decimal
        Dim ldecValorCrMaxCap = 0D, ldecValorCrMaxInt As Decimal, ldecValorCrMaxIva As Decimal
        Dim lobjFactura As New ClsFactura()
        Dim lobjValorLlave As Object(), lstrPref_Fact As String, lentIdFact As Integer
        Dim ldtbFactVivas = FdtbFacturasVivas()
        For Each ldrwFactViva As DataRow In ldtbFactVivas.Rows
            lstrPref_Fact = ClsPanorama.FobjValorCampo(ldrwFactViva(
                    ClsPrefijo_FactStr.SstrNombreCampoBd), EnuTipoValor.EnuString)
            lentIdFact = ClsPanorama.FobjValorCampo(ldrwFactViva(
                    ClsIdFacturaEnt.SstrNombreCampoBd), EnuTipoValor.EnuInteger)
            lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrPref_Fact, lentIdFact}
            lobjFactura.SAbra(lobjValorLlave)
            ldecValorCrMaxCap += lobjFactura.FdecDeudaCapitalAntesIva
            ldecValorCrMaxInt += lobjFactura.FdecDeudaIntMoraAntesIva
            ldecValorCrMaxIva += lobjFactura.FdecDeudaIva
        Next
        adecValorCrMaxInt = ldecValorCrMaxInt
        adecValorCrMaxIva = ldecValorCrMaxIva
        Return ldecValorCrMaxCap
    End Function
    Private Function FdtbFacturasVivas() As DataTable
        Dim lstrTabla = ClsFactura.SstrNombreTabla
        Dim lstrCampSel As String() = {ClsPrefijo_FactStr.SstrNombreCampoBd,
                ClsIdFacturaEnt.SstrNombreCampoBd, ClsFechaFacturaDtm.SstrNombreCampoBd}
        Dim lstrOrden As String(,) = {{ClsFechaFacturaDtm.SstrNombreCampoBd, "ASC"},
                {ClsIdFacturaEnt.SstrNombreCampoBd, "DESC"}}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdCliente_FactDbl.SstrNombreCampoBd & " = " & ObjIdCliente_NotaCrDbl.ObjValorPro &
                " AND " & ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd & " = '" &
                ObjIdPredioAgrupador_NotaCrStr.ToString() & "' AND (" &
                ClsDebitos_FactDec.SstrNombreCampoBd & " - " & ClsCreditos_FactDec.SstrNombreCampoBd &
                ") > 0"
        Dim ldtbRes = ClsPanorama.FdtbDataTable(lstrTabla, lstrCampSel, lstrOrden, lstrFiltro)
        Return ldtbRes
    End Function
    Friend Function FblnTipoDsctoValido()
        Dim lblnEsValido = False
        Select Case EnuTipoDsctoPorValor
            Case EnuTipoDescuentoDef.EnuDsctoCapital
                lblnEsValido = DecValorCrMaxCap > 0
            Case EnuTipoDescuentoDef.EnuDsctoIntMora
                lblnEsValido = DecValorCrMaxInt > 0
            Case EnuTipoDescuentoDef.EnuReteCree, EnuTipoDescuentoDef.EnuReteFuente,
                    EnuTipoDescuentoDef.EnuReteCree, EnuTipoDescuentoDef.EnuReteIca,
                    EnuTipoDescuentoDef.EnuReteIva
            Case EnuTipoDescuentoDef.EnuCancelaIva
                lblnEsValido = DecValorCrMaxIva > 0
        End Select
        Return lblnEsValido
    End Function
    Friend Function FblnValorDsctoValido()
        Dim lblnEsValido As Boolean
        Select Case EnuTipoDsctoPorValor
            Case EnuTipoDescuentoDef.EnuDsctoCapital
                lblnEsValido = MdecDsctoPorValor > 0 AndAlso MdecDsctoPorValor <=
                        DecValorCrMaxCap
            Case EnuTipoDescuentoDef.EnuDsctoIntMora
                lblnEsValido = MdecDsctoPorValor > 0 AndAlso MdecDsctoPorValor <=
                        DecValorCrMaxInt
            Case EnuTipoDescuentoDef.EnuCancelaIva
                lblnEsValido = MdecDsctoPorValor > 0 AndAlso MdecDsctoPorValor <=
                        DecValorCrMaxIva
            Case Else
                lblnEsValido = False
        End Select
        Return lblnEsValido
    End Function
    Friend Sub SApliqueDescuentoValor()
        Dim lobjFact As New ClsFactura, lstrPref As String, lentIdFac As Integer,
                lobjValorLlave As Object()
        Dim ldecValorPorAplicar = DecDsctoPorValor, ldecValorAplicado As Decimal
        Dim ldtbFactVivas = FdtbFacturasVivas()
        For Each ldrwFactViva As DataRow In ldtbFactVivas.Rows
            ldecValorAplicado = 0
            If ldecValorPorAplicar > 0 Then
                lstrPref = ClsPanorama.FobjValorCampo(ldrwFactViva(
                    ClsPrefijo_FactStr.SstrNombreCampoBd), EnuTipoValor.EnuString)
                lentIdFac = ClsPanorama.FobjValorCampo(ldrwFactViva(
                    ClsIdFacturaEnt.SstrNombreCampoBd), EnuTipoValor.EnuInteger)
                lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentIdFac}
                lobjFact.SAbra(lobjValorLlave)
                Select Case EnuTipoDsctoPorValor
                    Case EnuTipoDescuentoDef.EnuDsctoCapital
                        SApliqueDsctoCapitalFac(lobjFact, ldecValorPorAplicar, ldecValorAplicado)
                    Case EnuTipoDescuentoDef.EnuDsctoIntMora
                        SApliqueDsctoMoraFac(lobjFact, ldecValorPorAplicar, ldecValorAplicado)
                    Case EnuTipoDescuentoDef.EnuCancelaIva

                End Select
                ldecValorPorAplicar -= ldecValorAplicado
            Else
                Exit For
            End If
        Next
    End Sub
    Private Sub SApliqueDsctoCapitalFac(aobjFac As ClsFactura, adecValorPorAplicar As Decimal,
            ByRef adecValorAplicado As Decimal)
        Dim ldecValorAplicado As Decimal, ldecDeudaCapital As Decimal
        For Each lobjItemFac As ClsItemFactura In aobjFac.ColItemsFactura
            If adecValorPorAplicar > 0 Then
                ldecDeudaCapital = lobjItemFac.FdecDeudaCapital - lobjItemFac.FdecDeudaIva
                If adecValorPorAplicar >= ldecDeudaCapital Then
                    ldecValorAplicado = ldecDeudaCapital
                Else
                    ldecValorAplicado = adecValorPorAplicar
                End If
                Dim ldecBase = 0D
                Dim ldblTasa = ClsOrionCop.FdblTasaDscto(lobjItemFac,
                        EnuTipoDescuentoDef.EnuDsctoCapital, ldecValorAplicado, ldecBase)
                Dim lobjItemNotaCr = FobjNewItemNotaCr()
                lobjItemNotaCr.ObjPrefijoFact_ItemNotaCrStr.ObjValorPro =
                        aobjFac.ObjPrefijo_FactStr.ObjValorPro
                lobjItemNotaCr.ObjIdFactura_ItemNotaCrEnt.ObjValorPro =
                        aobjFac.ObjIdFacturaEnt.ObjValorPro
                lobjItemNotaCr.ObjIdItemFac_ItemNotaCrShr.ObjValorPro =
                        lobjItemFac.ObjIdItemFacturaShr.ObjValorPro
                lobjItemNotaCr.ObjBaseDscto_NotaCrDec.ObjValorPro = ldecBase
                lobjItemNotaCr.ObjValor_ItemNotaCrDec.ObjValorPro = ldecValorAplicado
                lobjItemNotaCr.ObjTasaDscto_ItemNotaCrDbl.ObjValorPro = ldblTasa
                lobjItemNotaCr.ObjEsReversionIvaBln.ObjValorPro = False
                lobjItemNotaCr.ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro =
                        EnuTipoDescuentoDef.EnuDsctoCapital
                SAdicioneNuevoItem(lobjItemNotaCr)
                adecValorAplicado += ldecValorAplicado
                adecValorPorAplicar -= ldecValorAplicado
            End If
        Next
    End Sub
    ' Descuento intereses de mora
    Private Sub SApliqueDsctoMoraFac(aobjFac As ClsFactura, adecValorPorAplicar As Decimal,
            ByRef adecValorAplicado As Decimal)
        Dim ldecValorAplicado As Decimal, ldecDeudaMora As Decimal
        For Each lobjItemFac As ClsItemFactura In aobjFac.ColItemsFactura
            If adecValorPorAplicar > 0 Then
                ldecDeudaMora = lobjItemFac.FdecDeudaIntMora
                If adecValorPorAplicar >= ldecDeudaMora Then
                    ldecValorAplicado = ldecDeudaMora
                Else
                    ldecValorAplicado = adecValorPorAplicar
                End If
                Dim ldecBase = 0D
                Dim ldblTasa = ClsOrionCop.FdblTasaDscto(lobjItemFac,
                        EnuTipoDescuentoDef.EnuDsctoIntMora, ldecValorAplicado, ldecBase)
                Dim lobjItemNotaCr = FobjNewItemNotaCr()
                lobjItemNotaCr.ObjPrefijoFact_ItemNotaCrStr.ObjValorPro =
                        aobjFac.ObjPrefijo_FactStr.ObjValorPro
                lobjItemNotaCr.ObjIdFactura_ItemNotaCrEnt.ObjValorPro =
                        aobjFac.ObjIdFacturaEnt.ObjValorPro
                lobjItemNotaCr.ObjIdItemFac_ItemNotaCrShr.ObjValorPro =
                        lobjItemFac.ObjIdItemFacturaShr.ObjValorPro
                lobjItemNotaCr.ObjBaseDscto_NotaCrDec.ObjValorPro = ldecBase
                lobjItemNotaCr.ObjValor_ItemNotaCrDec.ObjValorPro = ldecValorAplicado
                lobjItemNotaCr.ObjTasaDscto_ItemNotaCrDbl.ObjValorPro = ldblTasa
                lobjItemNotaCr.ObjEsReversionIvaBln.ObjValorPro = False
                lobjItemNotaCr.ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro =
                            EnuTipoDescuentoDef.EnuDsctoIntMora
                SAdicioneNuevoItem(lobjItemNotaCr)
                adecValorAplicado += ldecValorAplicado
                adecValorPorAplicar -= ldecValorAplicado
            End If
        Next
    End Sub
    ' Descuento a IVA (Iva al gasto)
    ' AVV Borrar?
    'Private Sub SApliqueDsctoIVA(aobjFac As ClsFactura, adecValorPorAplicar As Decimal,
    '        ByRef adecValorAplicado As Decimal)
    '    Dim ldecValorAplicado As Decimal, ldecDeudaIVA As Decimal
    '    For Each lobjItemFac As ClsItemFactura In aobjFac.ColItemsFactura
    '        If adecValorPorAplicar > 0 Then
    '            ' IVA Int Mora

    '            ldecDeudaIVA = lobjItemFac.FdecDeudaIvaInt
    '            If adecValorPorAplicar >= ldecDeudaIVA Then
    '                ldecValorAplicado = ldecDeudaIVA
    '            Else
    '                ldecValorAplicado = adecValorPorAplicar
    '            End If
    '            Dim ldecBase = 0D
    '            Dim ldblTasa = 0.0
    '            Dim lobjItemNotaCr = FobjNewItemNotaCr()
    '            lobjItemNotaCr.ObjPrefijoFact_ItemNotaCrStr.ObjValorPro =
    '                    aobjFac.ObjPrefijo_FactStr.ObjValorPro
    '            lobjItemNotaCr.ObjIdFactura_ItemNotaCrEnt.ObjValorPro =
    '                    aobjFac.ObjIdFacturaEnt.ObjValorPro
    '            lobjItemNotaCr.ObjIdItemFac_ItemNotaCrShr.ObjValorPro =
    '                    lobjItemFac.ObjIdItemFacturaShr.ObjValorPro
    '            lobjItemNotaCr.ObjBaseDscto_NotaCrDec.ObjValorPro = ldecBase
    '            lobjItemNotaCr.ObjValor_ItemNotaCrDec.ObjValorPro = ldecValorAplicado
    '            lobjItemNotaCr.ObjTasaDscto_ItemNotaCrDbl.ObjValorPro = ldblTasa
    '            lobjItemNotaCr.ObjEsReversionIvaBln.ObjValorPro = True
    '            lobjItemNotaCr.ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro =
    '                        EnuTipoDescuentoDef.EnuCancelaIva
    '            SAdicioneNuevoItem(lobjItemNotaCr)
    '            adecValorAplicado += ldecValorAplicado
    '            adecValorPorAplicar -= ldecValorAplicado
    '        End If
    '    Next
    '    For Each lobjItemFac As ClsItemFactura In aobjFac.ColItemsFactura
    '        If adecValorPorAplicar > 0 Then
    '            ' IVA Int Mora 
    '            ldecDeudaIVA = lobjItemFac.FdecDeudaIvaCapital
    '            If adecValorPorAplicar >= ldecDeudaIVA Then
    '                ldecValorAplicado = ldecDeudaIVA
    '            Else
    '                ldecValorAplicado = adecValorPorAplicar
    '            End If
    '            Dim ldecBase = 0D
    '            Dim ldblTasa = 0.0
    '            Dim lobjItemNotaCr = FobjNewItemNotaCr()
    '            lobjItemNotaCr.ObjPrefijoFact_ItemNotaCrStr.ObjValorPro =
    '                    aobjFac.ObjPrefijo_FactStr.ObjValorPro
    '            lobjItemNotaCr.ObjIdFactura_ItemNotaCrEnt.ObjValorPro =
    '                    aobjFac.ObjIdFacturaEnt.ObjValorPro
    '            lobjItemNotaCr.ObjIdItemFac_ItemNotaCrShr.ObjValorPro =
    '                    lobjItemFac.ObjIdItemFacturaShr.ObjValorPro
    '            lobjItemNotaCr.ObjBaseDscto_NotaCrDec.ObjValorPro = ldecBase
    '            lobjItemNotaCr.ObjValor_ItemNotaCrDec.ObjValorPro = ldecValorAplicado
    '            lobjItemNotaCr.ObjTasaDscto_ItemNotaCrDbl.ObjValorPro = ldblTasa
    '            lobjItemNotaCr.ObjEsReversionIvaBln.ObjValorPro = True
    '            lobjItemNotaCr.ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro =
    '                        EnuTipoDescuentoDef.EnuCancelaIva
    '            SAdicioneNuevoItem(lobjItemNotaCr)
    '            adecValorAplicado += ldecValorAplicado
    '            adecValorPorAplicar -= ldecValorAplicado
    '        End If
    '    Next
    'End Sub
#End Region
#Region "Procedimientos de Factura"
    Friend Sub SAnuleFactura(aobjFactura As ClsFactura)
        ' AVV Revisar por cambios en el mnejo de la deuda del IVA por intereses de mora
        ' que se está cargando al IVA total
        Dim lobjItemNotaCr As ClsItemNotaCr = Nothing
        Dim lstrPrefFact As String = aobjFactura.ObjPrefijo_FactStr.ObjValorPro
        Dim lentIdFac As Integer = aobjFactura.ObjIdFacturaEnt.ObjValorPro
        BlnAnulandoFac = True
        If GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro Then
            If Not Date.Today <= GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo Then
                Throw New ErrorInesperadoPanLException("No es posible anular esta factura por fuera del periodo actual!")
            End If
            ObjFecha_NotaCrDtm.ObjValorPro = Date.Today
        Else
            If Date.Today > GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo Then
                ObjFecha_NotaCrDtm.ObjValorPro =
                        GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
            Else
                ObjFecha_NotaCrDtm.ObjValorPro = Date.Today
            End If
        End If
        ObjIdCliente_NotaCrDbl.ObjValorPro = aobjFactura.ObjIdCliente_FactDbl.ObjValorPro
        ObjIdPredioAgrupador_NotaCrStr.ObjValorPro = aobjFactura.ObjIdPredioAgrupador_FacStr.ObjValorPro
        ObjComentario_NotaCrStr.ObjValorPro = "Anulación Factura Nro " & aobjFactura.StrNumeroFactura
        ObjModoNotaCrByt.ObjValorPro = EnuModoNotaCr.EnuPorFactura
        ObjIdTipoNotaCrByt.ObjValorPro = EnuTipoNotaCrDef.EnuAnulaFac
        For Each lobjItemFac As ClsItemFactura In aobjFactura.ColItemsFactura
            Dim ldecDeudaAntesIva = lobjItemFac.FdecDeudaCapital - lobjItemFac.FdecIva
            If ldecDeudaAntesIva > 0 Then
                lobjItemNotaCr = FobjNewItemNotaCr()
                With lobjItemNotaCr
                    .ObjEsReversionIvaBln.ObjValorPro = False
                    .ObjPrefijoFact_ItemNotaCrStr.ObjValorPro = lstrPrefFact
                    .ObjIdFactura_ItemNotaCrEnt.ObjValorPro = lentIdFac
                    .ObjIdItemFac_ItemNotaCrShr.ObjValorPro = lobjItemFac.ObjIdItemFacturaShr.ObjValorPro
                    .ObjBaseDscto_NotaCrDec.ObjValorPro = ldecDeudaAntesIva
                    .ObjTasaDscto_ItemNotaCrDbl.ObjValorPro = 1
                    .ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro = EnuTipoDescuentoDef.EnuDsctoCapital
                    .ObjValor_ItemNotaCrDec.ObjValorPro = ldecDeudaAntesIva
                End With
                ObjValor_NotaCrDec.ObjValorPro += lobjItemNotaCr.ObjValor_ItemNotaCrDec.ObjValorPro
                ColItemsNotaCr.Add(lobjItemNotaCr)
            End If
            If lobjItemFac.FdecIva > 0 Then
                lobjItemNotaCr = FobjNewItemNotaCr()
                With lobjItemNotaCr
                    .ObjEsReversionIvaBln.ObjValorPro = True
                    .ObjPrefijoFact_ItemNotaCrStr.ObjValorPro = lstrPrefFact
                    .ObjIdFactura_ItemNotaCrEnt.ObjValorPro = lentIdFac
                    .ObjIdItemFac_ItemNotaCrShr.ObjValorPro =
                            lobjItemFac.ObjIdItemFacturaShr.ObjValorPro
                    .ObjBaseDscto_NotaCrDec.ObjValorPro = lobjItemFac.FdecBaseIva
                    .ObjTasaDscto_ItemNotaCrDbl.ObjValorPro = lobjItemFac.FdblTasaIva
                    .ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro = EnuTipoDescuentoDef.EnuDsctoCapital
                    .ObjValor_ItemNotaCrDec.ObjValorPro = lobjItemFac.FdecIva
                End With
                ObjValor_NotaCrDec.ObjValorPro += lobjItemNotaCr.ObjValor_ItemNotaCrDec.ObjValorPro
                ColItemsNotaCr.Add(lobjItemNotaCr)
            End If
            If lobjItemFac.FdecDeudaIntMora > 0 Then
                lobjItemNotaCr = FobjNewItemNotaCr()
                With lobjItemNotaCr
                    .ObjEsReversionIvaBln.ObjValorPro = False
                    .ObjPrefijoFact_ItemNotaCrStr.ObjValorPro = lstrPrefFact
                    .ObjIdFactura_ItemNotaCrEnt.ObjValorPro = lentIdFac
                    .ObjIdItemFac_ItemNotaCrShr.ObjValorPro = lobjItemFac.ObjIdItemFacturaShr.
                                    ObjValorPro
                    .ObjBaseDscto_NotaCrDec.ObjValorPro = lobjItemFac.FdecDeudaIntMora
                    .ObjTasaDscto_ItemNotaCrDbl.ObjValorPro = 1
                    .ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro = EnuTipoDescuentoDef.
                                EnuDsctoIntMora
                    .ObjValor_ItemNotaCrDec.ObjValorPro = lobjItemFac.FdecDeudaIntMora
                End With
                ObjValor_NotaCrDec.ObjValorPro += lobjItemNotaCr.ObjValor_ItemNotaCrDec.
                            ObjValorPro
                ColItemsNotaCr.Add(lobjItemNotaCr)
                'Dim ldecIvaIntMora = lobjItemFac.FdecDeudaIvaInt
                'If ldecIvaIntMora > 0 Then
                '    lobjItemNotaCr = FobjNewItemNotaCr()
                '    With lobjItemNotaCr
                '        .ObjEsReversionIvaBln.ObjValorPro = True
                '        .ObjPrefijoFact_ItemNotaCrStr.ObjValorPro = lstrPrefFact
                '        .ObjIdFactura_ItemNotaCrEnt.ObjValorPro = lentIdFac
                '        .ObjIdItemFac_ItemNotaCrShr.ObjValorPro =
                '            lobjItemFac.ObjIdItemFacturaShr.ObjValorPro
                '        .ObjBaseDscto_NotaCrDec.ObjValorPro = lobjItemFac.FdecDeudaIntMora
                '        .ObjTasaDscto_ItemNotaCrDbl.ObjValorPro = lobjItemFac.FdblTasaIva
                '        .ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro = EnuTipoDescuentoDef.
                '                EnuDsctoIntMora
                '        .ObjValor_ItemNotaCrDec.ObjValorPro = ldecIvaIntMora
                '    End With
                '    ObjValor_NotaCrDec.ObjValorPro += lobjItemNotaCr.ObjValor_ItemNotaCrDec.
                '            ObjValorPro
                '    ColItemsNotaCr.Add(lobjItemNotaCr)
                'End If
            End If
        Next
        SActualice(True)
        aobjFactura.SAnuleFactura(Me)
    End Sub
    ''' <summary>
    ''' Indica si una nota crédito esta haciendo una devolución total del servicio facturado
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>Aplica fara notas cr cuando esta instalado el módulo de facturación electrónica</remarks>
    Private Function FblnEsDesctoTotal() As Boolean
        Dim lblnEsDsctoTotal = False
        Dim lobjFra As New ClsFactura()
        Dim lstrPrefFac As String, lentIdFac As Integer
        Dim lobjItemNcr As ClsItemNotaCr = ColItemsNotaCr("1")
        lstrPrefFac = lobjItemNcr.ObjPrefijoFact_ItemNotaCrStr.ObjValorPro
        lentIdFac = lobjItemNcr.ObjIdFactura_ItemNotaCrEnt.ObjValorPro
        Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, lstrPrefFac, lentIdFac}
        lobjFra.SAbra(lobjValorLlave)
        If lobjFra.ObjDebitos_FactDec.ObjValorPro = lobjFra.ObjCreditos_FactDec.ObjValorPro Then
            lblnEsDsctoTotal = (lobjFra.ObjDebitos_FactDec.ObjValorPro = ObjValor_NotaCrDec.ObjValorPro)
        End If
        Return lblnEsDsctoTotal
    End Function
#End Region
#Region "Manejo Items Nota Cr"
    Friend ReadOnly Property ColItemsNotaCr As Collection
        Get
            If IsNothing(McolItemsNotaCr) Then
                McolItemsNotaCr = New Collection
            End If
            If ObjIdNotaCrEnt.BlnEsValido AndAlso McolItemsNotaCr.Count = 0 AndAlso
                    EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                Dim ldtbItems = FdtbItemsNotaCr(False)
                If Not IsNothing(ldtbItems) Then
                    For Each ldrwItemNotaCr As DataRow In ldtbItems.Rows
                        Dim lobjItemNotaCr As New ClsItemNotaCr(Me, ldrwItemNotaCr)
                        lobjItemNotaCr.SLeaValores(True)
                        McolItemsNotaCr.Add(lobjItemNotaCr, lobjItemNotaCr.ObjIdItemNotaCrShr.ToString)
                    Next
                End If
            End If
            Return McolItemsNotaCr
        End Get
    End Property
    Friend Function FobjNewItemNotaCr() As ClsItemNotaCr
        Dim lobjItemNotaCr As ClsItemNotaCr = Nothing
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            Dim ldtbItemsNCr = FdtbItemsNotaCr(False)
            Dim ldrwItemNotaCr As DataRow = ldtbItemsNCr.NewRow
            lobjItemNotaCr = New ClsItemNotaCr(Me, ldrwItemNotaCr)
            With lobjItemNotaCr
                .SCreeObj(Nothing)
            End With
        End If
        Return lobjItemNotaCr
    End Function
    Friend Sub SAdicioneNuevoItem(aobjNuevoItemNotaCr As ClsItemNotaCr)
        Try
            If IsNothing(McolItemsNotaCr) Then
                McolItemsNotaCr = ColItemsNotaCr
            End If
            McolItemsNotaCr.Add(aobjNuevoItemNotaCr)
            SEstablezcaValorNotaCr()
            ObjFecha_NotaCrDtm.SValideFechaConItems()
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
    End Sub
    Friend Sub SElimineItemNotaCr(ashrIdItemNotaCr As Short)
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            If McolItemsNotaCr.Count > 0 AndAlso McolItemsNotaCr.Count >= ashrIdItemNotaCr Then
                McolItemsNotaCr.Remove(ashrIdItemNotaCr)
            Else
                Throw New ErrorInesperadoPanLException("Objeto item nota cr no encontrado!")
            End If
            SEstablezcaValorNotaCr()
        End If
    End Sub
    Friend Function FblnExisteItemNotaCr(astrPrefFactura As String, aentIdFactura As Integer,
            ashrIdItemFac As Short, aenuTipoDscto As EnuTipoDescuentoDef,
            ablnEsReversionIva As Boolean) As Boolean
        Dim lblnExiste = False
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando AndAlso Not IsNothing(ColItemsNotaCr) Then
            Dim lstrPrefFact = String.Empty, lentIdFact = 0, lshrIdItemFac = 0S, lblnEsRevIva = False
            Dim lenuTipoDscto = EnuTipoDescuentoDef.None
            For Each lobjItemNotaCr As ClsItemNotaCr In ColItemsNotaCr
                With lobjItemNotaCr
                    lstrPrefFact = .ObjPrefijoFact_ItemNotaCrStr.ObjValorPro
                    lentIdFact = .ObjIdFactura_ItemNotaCrEnt.ObjValorPro
                    lshrIdItemFac = .ObjIdItemFac_ItemNotaCrShr.ObjValorPro
                    lenuTipoDscto = .ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro
                    lblnEsRevIva = .ObjEsReversionIvaBln.ObjValorPro
                    lblnExiste = (lstrPrefFact = astrPrefFactura AndAlso lentIdFact = aentIdFactura AndAlso
                                  lshrIdItemFac = ashrIdItemFac AndAlso aenuTipoDscto = lenuTipoDscto AndAlso
                                  ablnEsReversionIva = lblnEsRevIva)
                    If lblnExiste Then Exit For
                End With
            Next
        End If
        Return lblnExiste
    End Function
    Friend Function FblnExisteNuevoItem(astrPrefFact As String, aentIdFact As Integer,
            ashrIdItemFac As Short, aenuTipoDscto As EnuTipoDescuentoDef) As Boolean
        Dim lstrPrefFact = String.Empty, lentIdFact = 0, lshrIdItemFac = 0S, lblnExiste = False
        Dim lenuTipoDscto = EnuTipoDescuentoDef.None
        For Each lobjItemNotaCr As ClsItemNotaCr In ColItemsNotaCr
            With lobjItemNotaCr
                lstrPrefFact = .ObjPrefijoFact_ItemNotaCrStr.ObjValorPro
                lentIdFact = .ObjIdFactura_ItemNotaCrEnt.ObjValorPro
                lshrIdItemFac = .ObjIdItemFac_ItemNotaCrShr.ObjValorPro
                lenuTipoDscto = .ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro
                lblnExiste = (lstrPrefFact = astrPrefFact AndAlso lentIdFact = aentIdFact AndAlso
                                  lshrIdItemFac = ashrIdItemFac AndAlso aenuTipoDscto = lenuTipoDscto)
                If lblnExiste Then
                    lblnExiste = Not (BlnAnulandoFac AndAlso (aenuTipoDscto =
                            EnuTipoDescuentoDef.EnuDsctoCapital OrElse aenuTipoDscto =
                            EnuTipoDescuentoDef.EnuDsctoIntMora))
                End If
                If lblnExiste Then Exit For
            End With
        Next
        Return lblnExiste
    End Function
    Friend Function FdtbItemsNotaCr(ablnRelItemFac As Boolean) As DataTable
        Dim ldtbItemsNotaCr As DataTable
        If ablnRelItemFac Then
            Dim lstrExpSqlItemNCr = FstrExpSqlItems()
            Dim lstrExpSql = "SELECT CONCAT(" & ClsPrefijoFact_ItemNotaCrStr.SstrNombreCampoBd &
                    ",'-', R." & ClsIdFactura_ItemNotaCrEnt.SstrNombreCampoBd &
                    ") AS NroFac, " & ClsDetalle_ItemFactStr.SstrNombreCampoBd &
                    ", Dato ," & ClsBaseDsctoDec.SstrNombreCampoBd & ", " &
                    ClsTasaDscto_ItemNotaCrDbl.SstrNombreCampoBd & ", R." &
                    ClsValor_ItemNotaCrDec.SstrNombreCampoBd & " FROM (" &
                    lstrExpSqlItemNCr & ") AS R INNER JOIN " &
                    ClsItemFactura.SstrNombreTabla & " AS F ON R." &
                    ClsIdCarpetaShr.SstrNombreCampoBd & " = F." &
                    ClsIdCarpetaShr.SstrNombreCampoBd & " AND R." &
                    ClsIdCentroUtilShr.SstrNombreCampoBd & " = F." &
                    ClsIdCentroUtilShr.SstrNombreCampoBd & " AND R." &
                    ClsPrefijoFact_ItemNotaCrStr.SstrNombreCampoBd & " = F." &
                    ClsPrefijo_ItemFactStr.SstrNombreCampoBd & " AND R." &
                    ClsIdFactura_ItemNotaCrEnt.SstrNombreCampoBd & " = F." &
                    ClsIdFactura_ItemFactEnt.SstrNombreCampoBd & " AND R." &
                    ClsIdItemFac_ItemNotaCrShr.SstrNombreCampoBd & " = F." &
                    ClsIdItemFacturaShr.SstrNombreCampoBd & " WHERE F." &
                    ClsIdCarpetaShr.SstrNombreCampoBd & " = " & GshrIdCarpeta & " AND F." &
                    ClsIdCentroUtilShr.SstrNombreCampoBd & " = " & GshrIdCentroUtil
            ldtbItemsNotaCr = ClsPanorama.FdtbDataTable(lstrExpSql)
        Else
            Dim lstrPref = "", lstrIdNotaCr = "0"
            If BlnExiste Then
                lstrPref = ObjPrefijo_NotaCrStr.ToString()
                lstrIdNotaCr = ObjIdNotaCrEnt.ToString()
            End If
            Dim lstrTabla = ClsItemNotaCr.SstrNombreTabla
            Dim lstrCampSel As String() = {"*"}
            Dim lstrOrden As String(,) = {{ClsIdItemNotaCrShr.SstrNombreCampoBd, "ASC"}}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsPrefijo_NotaCrStr.SstrNombreCampoBd & " = '" & lstrPref &
                    "' AND " & ClsIdNotaCr_ItemNotaCrEnt.SstrNombreCampoBd & " = " &
                    lstrIdNotaCr

            ldtbItemsNotaCr = ClsPanorama.FdtbDataTable(lstrTabla, lstrCampSel,
                    lstrOrden, lstrFiltro)
        End If
        Return ldtbItemsNotaCr
    End Function
    Friend Function FdecValorDescuento(aenuTipoDescuento As EnuTipoDescuentoDef) As Decimal
        Dim ldecDscto = 0D
        For Each lobjitemNCr As ClsItemNotaCr In ColItemsNotaCr
            If lobjitemNCr.ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro = aenuTipoDescuento Then
                ldecDscto += lobjitemNCr.ObjValor_ItemNotaCrDec.ObjValorPro
            End If
        Next
        Return ldecDscto
    End Function
    Private Function FstrExpSqlItems() As String
        Dim lstrIdNotaCr As String = ObjIdNotaCrEnt.ToString()
        If String.IsNullOrEmpty(lstrIdNotaCr) Then
            lstrIdNotaCr = "0"
        End If
        Dim lstrExpSql = "SELECT " & ClsIdCarpetaShr.SstrNombreCampoBd & ", " &
                ClsIdCentroUtilShr.SstrNombreCampoBd & ", " &
                ClsPrefijoFact_ItemNotaCrStr.SstrNombreCampoBd & ", " &
                ClsIdFactura_ItemNotaCrEnt.SstrNombreCampoBd & ", " &
                ClsIdItemFac_ItemNotaCrShr.SstrNombreCampoBd & ", Dato, " &
                ClsBaseDscto_NotaCrDec.SstrNombreCampoBd & ", " &
                ClsTasaDscto_ItemNotaCrDbl.SstrNombreCampoBd & ", " &
                ClsValor_ItemNotaCrDec.SstrNombreCampoBd & " FROM " &
                ClsItemNotaCr.SstrNombreTabla & " AS I INNER JOIN oritblconstantes" &
                " AS C ON C.idgrupo = " & EnuGrupoConstantesOriDef.EnuTipoDescuento &
                " AND I." & ClsIdTipoDscto_ItemNotaCrByt.SstrNombreCampoBd & " = " &
                "C.IdConstante WHERE " & ClsPrefijo_NotaCrStr.SstrNombreCampoBd &
                " = '" & ObjPrefijo_NotaCrStr.ToString() & "' AND " &
                ClsIdNotaCrEnt.SstrNombreCampoBd & " = " & lstrIdNotaCr
        Return lstrExpSql
    End Function
    Private Sub SEstablezcaValorNotaCr()
        Dim ldecValorNotaCr = 0D
        If Not IsNothing(McolItemsNotaCr) AndAlso McolItemsNotaCr.Count > 0 Then
            For Each lobjItemNotaCr As ClsItemNotaCr In McolItemsNotaCr
                ldecValorNotaCr += lobjItemNotaCr.ObjValor_ItemNotaCrDec.ObjValorPro
            Next
        End If
        ObjValor_NotaCrDec.ObjValorPro = ldecValorNotaCr
    End Sub
#End Region
#Region "Nota RCr"
    Private Sub SGenereNotaRevCr()
        Dim lstrPref = GobjParametros.FstrPrefijoDoc(EnuIdDocumentoDef.EnuNotaReversaCr)
        Dim lobjNotaRcr As New ClsNotaReversionCr(lstrPref)
        lobjNotaRcr.SCreeObj(Nothing)
        With lobjNotaRcr
            Dim ldtmFechaNotaRcr As Date = GCDTMFECHANULA
            Dim ldtmFechaAnu As Date = ObjFechaAnulacion_NotaCrDtm.ObjValorPro
            If IsDate(ldtmFechaAnu) Then
                ldtmFechaNotaRcr = ldtmFechaAnu.Date
            Else
                Throw New ErrorInesperadoPanLException("Se esperaba una fecha!")
            End If
            .ObjTipoDocReversadoByt.ObjValorPro = EnuDocReversado.EnuNotaCr
            .ObjFecha_NotaReversaCrDtm.ObjValorPro = ldtmFechaNotaRcr
            .ObjDetalle_NotaReversaCrStr.ObjValorPro = "Anulacion Nota Crédito " & StrNumeroNotaCr
            .ObjIdPredioAgrupador_NotaReversaCrStr.ObjValorPro = ObjIdPredioAgrupador_NotaCrStr.ObjValorPro
            .ObjPrefijoDoc_NotaReversaCrStr.ObjValorPro = ObjPrefijo_NotaCrStr.ObjValorPro
            .ObjIdDoc_NotaReversaCrEnt.ObjValorPro = ObjIdNotaCrEnt.ObjValorPro
            .ObjIdCliente_NotaReversaCrDbl.ObjValorPro = ObjIdCliente_NotaCrDbl.ObjValorPro
            .ObjValor_NotaReversaCrDec.ObjValorPro = ObjValor_NotaCrDec.ObjValorPro
            .ObjDocReversado = Me
            .SActualice(True)
        End With
    End Sub
    Friend ReadOnly Property StrNroNotaRCr As String
        Get
            If String.IsNullOrEmpty(MstrNroNotaRCr) Then
                Dim lentIdNotaCr = 0, lstrPrefNotaCr = String.Empty
                If Not IsNothing(ObjIdNotaCrEnt) Then
                    lentIdNotaCr = ObjIdNotaCrEnt.ObjValorPro
                End If
                If Not IsNothing(ObjPrefijo_NotaCrStr) Then
                    lstrPrefNotaCr = ObjPrefijo_NotaCrStr.ObjValorPro
                End If
                Dim lstrCamposSelect As String() = {ClsPrefijo_NotaReversaCrStr.SstrNombreCampoBd,
                        ClsIdNotaReversaCrEnt.SstrNombreCampoBd}
                Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                        ClsTipoDocReversadoByt.SstrNombreCampoBd & " = " &
                        EnuDocReversado.EnuNotaCr & " AND " &
                        ClsPrefijoDoc_NotaReversaCrStr.SstrNombreCampoBd & " = '" & lstrPrefNotaCr &
                        "' AND " & ClsIdDoc_NotaReversaCrEnt.SstrNombreCampoBd & " = " &
                        lentIdNotaCr
                Dim ldtbNotasRCr = ClsPanorama.FdtbDataTable(ClsNotaReversionCr.SstrNombreTabla,
                        lstrCamposSelect, {{}}, lstrFiltro)
                If ldtbNotasRCr.Rows.Count > 0 Then
                    Dim lstrPref As String = ClsPanorama.FobjValorCampo(ldtbNotasRCr(0)(0),
                            EnuTipoValor.enuString)
                    Dim lentId As Integer = ClsPanorama.FobjValorCampo(ldtbNotasRCr(0)(1),
                            EnuTipoValor.enuInteger)
                    MstrNroNotaRCr = ClsPanorama.FstrNumeroDcto(lstrPref, lentId)
                End If
            End If
            Return MstrNroNotaRCr
        End Get
    End Property
    Friend ReadOnly Property ObjNotaReversionCr As ClsNotaReversionCr
        Get
            If IsNothing(MobjNotaReversionRC) AndAlso Not String.IsNullOrEmpty(StrNroNotaRCr) Then
                Dim lstrPref As String = ClsPanorama.FstrPrefijoDcto(StrNroNotaRCr)
                Dim lentId As Integer = ClsPanorama.FentIdDcto(StrNroNotaRCr)
                Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentId}
                MobjNotaReversionRC = New ClsNotaReversionCr()
                MobjNotaReversionRC.SAbra(lobjValorLlave)
            End If
            Return MobjNotaReversionRC
        End Get
    End Property
#End Region
#Region "eFac"
    ''' <summary>
    ''' Indica si es una factura electrónica y si esta registrada 
    ''' </summary>
    Friend ReadOnly Property BlnEstaRegEFac As Boolean
        Get
            Return BlnEsDocEle AndAlso ObjIdEstadoEDocEnt.ObjValorPro >= EnuEstadoEDoc.EnuRegi
        End Get
    End Property
    Friend ReadOnly Property BlnEsDocEle As Boolean
        Get
            Dim lblnEsDocEle = GobjParametros.BlnEFacAutorizado AndAlso
                    ObjIdEstadoEDocEnt.ObjValorPro <> EnuEstadoEDoc.EnuNoEDoc
            Return lblnEsDocEle
        End Get
    End Property
    Public Function BlnAfectaFrasRegEFac() As Boolean
        Dim lblnHay = Not BlnEsRetencion
        If lblnHay Then
            lblnHay = False
            For Each lobjItemNCr As ClsItemNotaCr In ColItemsNotaCr
                If lobjItemNCr.BlnEsDscto Then
                    lblnHay = lobjItemNCr.ObjFactura.BlnEsFacEle AndAlso
                            Not lobjItemNCr.ObjFactura.FenuVerFacEFac = EnuVerEFac.EnuV1
                    If lblnHay Then Exit For
                End If
            Next
        End If
        Return lblnHay
    End Function
    Friend Function BlnAfectaEFacV1() As Boolean
        Dim lblnAfecta As Boolean
        lblnAfecta = (ObjIdTipoNotaCrByt.ObjValorPro = EnuTipoNotaCrDef.EnuDescuento) OrElse
                ObjIdTipoNotaCrByt.ObjValorPro = EnuTipoNotaCrDef.EnuAnulaFac
        If lblnAfecta Then
            lblnAfecta = False
            For Each lobjItemNCr As ClsItemNotaCr In ColItemsNotaCr
                If lobjItemNCr.BlnEsDscto Then
                    lblnAfecta = (lobjItemNCr.ObjFactura.FenuVerFacEFac = EnuVerEFac.EnuV1)
                    If lblnAfecta Then Exit For
                End If
            Next
        End If
        Return lblnAfecta
    End Function
    Friend Function FblnInsertarEFac() As Boolean
        Dim lblnReg = False
        If GobjParametros.BlnEFacAutorizado AndAlso BlnAfectaFrasRegEFac() Then
            lblnReg = (ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuNoReg) AndAlso
                    BlnEsDocEle
            If lblnReg Then
                For Each lobjItemNCr As ClsItemNotaCr In ColItemsNotaCr
                    lblnReg = lobjItemNCr.ObjFactura.BlnEsFacEle AndAlso
                            (Not lobjItemNCr.ObjFactura.FenuVerFacEFac = EnuVerEFac.EnuV1) AndAlso
                            lobjItemNCr.ObjFactura.ObjIdEstadoEDocEnt.ObjValorPro >=
                            EnuEstadoEDoc.EnuRegi AndAlso
                            lobjItemNCr.ObjFactura.ObjIdEstadoEDocEnt.ObjValorPro <>
                            EnuEstadoEDoc.EnuNoEDoc
                    If lblnReg Then Exit For
                Next
            End If
        End If
        Return lblnReg
    End Function
    Friend Function FblnActualizarEstEFac() As Boolean
        Dim lblnActu = False
        If BlnEsDocEle AndAlso BlnAfectaFrasRegEFac() Then
            lblnActu = (ObjIdEstadoEDocEnt.ObjValorPro < EnuEstadoEDoc.EnuEnviada) AndAlso
                    (ObjIdEstadoEDocEnt.ObjValorPro >= EnuEstadoEDoc.EnuEnProceso)
        End If
        Return lblnActu
    End Function
    Friend Function FblnEnviarEFac() As Boolean
        Dim lblnEnv = False
        If BlnEsDocEle AndAlso BlnAfectaFrasRegEFac() Then
            lblnEnv = (ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuRegi)
        End If
        Return lblnEnv
    End Function
    Friend Sub SHabiliteProcesarEFac()
        If GobjParametros.BlnEFacAutorizado AndAlso BlnAfectaFrasRegEFac() Then
            If ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuInvalida OrElse
                    ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuErrorFtp Then
                If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                    EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                End If
                ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuNoReg
                SActualice(True)
            End If
        End If
    End Sub
    ''' <summary>
    ''' Devuelve un array donde cada elemento contiene el prefijo de la factura, el id de la factura,
    ''' el item de factura con descuento, el tipo de descuento y el valor del descuento.
    ''' separados por ampersand
    ''' </summary>
    ''' <returns></returns>
    Friend Function FstrDctos(ablnReversion As Boolean) As String()
        Dim lstrDscto As String, i = -1
        Dim lstrDsctosNcr As String() = FstrDsctosNcr()
        If lstrDsctosNcr.Length > 0 Then
            Dim lstrDsctos(UBound(lstrDsctosNcr)) As String
            For Each lstrDsctoNcr As String In lstrDsctosNcr
                i += 1
                lstrDscto = FstrValorDscto(lstrDsctoNcr, ablnReversion)
                lstrDsctos(i) = lstrDsctoNcr & "&" & lstrDscto
            Next
            Return lstrDsctos
        End If
        Return Array.Empty(Of String)
    End Function
    Private Function FstrDsctosNcr() As String()
        Dim lobjFactura As New ClsFactura()
        Dim lobjValorLlave As Object()
        Dim lstrPref As String, lentIdFac As Integer
        Dim lenuTipoDscto As EnuTipoDescuentoDef
        Dim lstrDsctos As String() = Array.Empty(Of String), lstrDscto As String
        Dim lstrTipoDes As String, i = -1
        For Each lobjItemNcr As ClsItemNotaCr In ColItemsNotaCr
            lstrPref = lobjItemNcr.ObjPrefijoFact_ItemNotaCrStr.ObjValorPro
            lentIdFac = lobjItemNcr.ObjIdFactura_ItemNotaCrEnt.ObjValorPro
            lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentIdFac}
            lobjFactura.SAbra(lobjValorLlave)
            If lobjFactura.ObjIdEstadoEDocEnt.ObjValorPro >= EnuEstadoEDoc.EnuRegi Then
                If lobjItemNcr.FblnEsDscto Then
                    lenuTipoDscto = lobjItemNcr.ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro
                    If lenuTipoDscto = EnuTipoDescuentoDef.EnuDsctoPP Then
                        lenuTipoDscto = EnuTipoDescuentoDef.EnuDsctoCapital
                    End If
                    lstrTipoDes = CType(lenuTipoDscto, Byte).ToString
                    lstrDscto = lstrPref & "&" & lobjItemNcr.ObjIdFactura_ItemNotaCrEnt.ToString & "&" &
                            lobjItemNcr.ObjIdItemFac_ItemNotaCrShr.ToString & "&" & lstrTipoDes
                    If Not lstrDsctos.Contains(lstrDscto) Then
                        i += 1
                        ReDim Preserve lstrDsctos(i)
                        lstrDsctos(i) = lstrDscto
                    End If
                End If
            End If
        Next
        Return lstrDsctos
    End Function
    Private Function FstrValorDscto(astrDsctoNcr As String, ablnReversion As Boolean) As String
        Dim lstrPartes As String() = astrDsctoNcr.Split("&")
        Dim ldecVlrDscto = 0D, lstrVlrDscto = String.Empty, lstrTipoItem As String
        Dim lenuTipoDscto As EnuTipoDescuentoDef
        For Each lobjItemNcr As ClsItemNotaCr In ColItemsNotaCr
            lenuTipoDscto = lobjItemNcr.ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro
            If lenuTipoDscto = EnuTipoDescuentoDef.EnuDsctoPP Then
                lenuTipoDscto = EnuTipoDescuentoDef.EnuDsctoCapital
            End If
            lstrTipoItem = CType(lenuTipoDscto, Byte).ToString
            If lobjItemNcr.ObjPrefijoFact_ItemNotaCrStr.ToString = lstrPartes(0) AndAlso
                    lobjItemNcr.ObjIdFactura_ItemNotaCrEnt.ToString = lstrPartes(1) AndAlso
                    lobjItemNcr.ObjIdItemFac_ItemNotaCrShr.ToString = lstrPartes(2) AndAlso
                    lstrTipoItem = lstrPartes(3) Then
                If Not ablnReversion Then
                    ldecVlrDscto += lobjItemNcr.ObjValor_ItemNotaCrDec.ObjValorPro
                Else
                    ldecVlrDscto += FdecValorDsctoReversado(lobjItemNcr.ObjIdItemNotaCrShr.ObjValorPro)
                End If
            End If
        Next
        If (ldecVlrDscto) > 0 Then
            lstrVlrDscto = Format(ldecVlrDscto, "#0.00")
        End If
        Return lstrVlrDscto
    End Function
    Private Function FdecValorDsctoReversado(ashrIdItemNcr As Short) As Decimal
        Dim ldecVlr = 0D
        Dim lobjItemNcr As ClsItemNotaCr = McolItemsNotaCr(ashrIdItemNcr.ToString)
        Dim lenuTipoNov As EnuTipoNov = lobjItemNcr.FenuTipoNovedad
        Dim lenuTipoNovRever As EnuTipoNov = ClsOrionCop.FenuTipoNovContraria(lenuTipoNov)
        For Each lobjNov As ClsNovedad In ColNovedades
            If lobjNov.ObjIdTipoNovedadByt.ObjValorPro = lenuTipoNovRever AndAlso
                    lobjNov.ObjIdItemDocOrigen_NovShr.ObjValorPro = ashrIdItemNcr Then
                ldecVlr += lobjNov.ObjValor_NovDec.ObjValorPro
            End If
        Next
        Return ldecVlr
    End Function
    Friend Function FstrRetenciones(ablnReversion As Boolean,
            aenuTipoProveEFac As EnuProveedorEFac) As String()
        Dim lstrRetenciones As String() = Array.Empty(Of String), i = -1
        Dim lstrRetensTipo As String()
        Dim lenuTipoRet As EnuTipoDescuentoDef = EnuTipoDescuentoDef.EnuReteFuente
        Do While lenuTipoRet <= EnuTipoDescuentoDef.EnuReteIva
            lstrRetensTipo = FstrRetenTipo(lenuTipoRet, ablnReversion, aenuTipoProveEFac)
            If lstrRetensTipo.Length > 0 Then
                For Each lstrRetenTipo As String In lstrRetensTipo
                    i += 1
                    ReDim Preserve lstrRetenciones(i)
                    lstrRetenciones(i) = lstrRetenTipo
                Next
            End If
            lenuTipoRet += 1
        Loop
        Return lstrRetenciones
    End Function
    Private Function FstrRetenTipo(aenuTipoRet As EnuTipoDescuentoDef, ablnReversion As Boolean,
            aenuTipoProveEFac As EnuProveedorEFac) As String()
        Dim lstrTasa As String, lstrTipoRetTasa As String, ldecRet As Decimal
        Dim lstrRet As String, ldecBase = 0D, lstrBase As String
        Dim i = -1
        Dim lstrRetenTipo As String() = Array.Empty(Of String)
        Dim lstrTiposRetTasa As String() = Array.Empty(Of String)
        Dim ldblTasas As Double() = FdblTasasRetRec(aenuTipoRet)
        Dim lstrTipoRet = FstrTipoRetencion(aenuTipoRet, aenuTipoProveEFac)
        For Each ldblTasa As Double In ldblTasas
            lstrTasa = Format(ldblTasa * 100, "#0.00")
            lstrTipoRetTasa = lstrTipoRet & "&" & lstrTasa
            If Not lstrTiposRetTasa.Contains(lstrTipoRetTasa) Then
                i += 1
                ReDim Preserve lstrTiposRetTasa(i)
                lstrTiposRetTasa(i) = lstrTipoRetTasa
                ldecRet = FdecVlrRetencion(aenuTipoRet, ldblTasa, ablnReversion, ldecBase)
                lstrRet = Format(ldecRet, "#0.00")
                lstrBase = Format(ldecBase, "#0.00")
                ReDim Preserve lstrRetenTipo(i)
                lstrRetenTipo(i) = lstrTipoRetTasa & "&" & lstrRet & "&" & lstrBase
            End If
        Next
        Return lstrRetenTipo
    End Function
    Private Function FdblTasasRetRec(aenuTipoRet As EnuTipoDescuentoDef) As Double()
        Dim ldblTasas As Double() = Array.Empty(Of Double), ldblTasa As Double, i = -1
        Dim lenuTipoDsctoItem As EnuTipoDescuentoDef
        For Each lobjItemNcr As ClsItemNotaCr In ColItemsNotaCr
            If lobjItemNcr.FblnEsRetencion Then
                lenuTipoDsctoItem = lobjItemNcr.ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro
                If lenuTipoDsctoItem = aenuTipoRet Then
                    ldblTasa = lobjItemNcr.ObjTasaDscto_ItemNotaCrDbl.ObjValorPro
                    If Not ldblTasas.Contains(ldblTasa) Then
                        i += 1
                        ReDim Preserve ldblTasas(i)
                        ldblTasas(i) = ldblTasa
                    End If
                End If
            End If
        Next
        Return ldblTasas
    End Function
    Private Function FdecVlrRetencion(aenuTipoReten As EnuTipoDescuentoDef,
            adblTasa As Double, ablnReversion As Boolean, ByRef adecVlrBase As Decimal)
        Dim ldecVlrRet = 0D, ldecBase = 0D, ldblTasaNov As Double
        Dim lenuTipoNov As EnuTipoNov = FenuTipoNov(aenuTipoReten)
        adblTasa = Math.Round(adblTasa, 4)
        If ablnReversion Then
            lenuTipoNov = ClsOrionCop.FenuTipoNovContraria(lenuTipoNov)
        End If
        For Each lobjNov As ClsNovedad In ColNovedades
            ldblTasaNov = Math.Round(lobjNov.ObjFactorDbl.ObjValorPro, 4)
            If lobjNov.ObjIdTipoNovedadByt.ObjValorPro = lenuTipoNov AndAlso
                                ldblTasaNov = adblTasa Then
                ldecVlrRet += lobjNov.ObjValor_NovDec.ObjValorPro
                ldecBase += lobjNov.ObjBaseDec.ObjValorPro
            End If
        Next
        adecVlrBase = ldecBase
        Return ldecVlrRet
    End Function
    Private Shared Function FstrTipoRetencion(aenuTipoRetencion As EnuTipoDescuentoDef,
            aenuProveedorEFactura As EnuProveedorEFac) As String
        Dim lstrTipoImp = "0"
        If aenuProveedorEFactura = EnuProveedorEFac.EnuProtecdataMisFac Then
            Select Case aenuTipoRetencion
                Case EnuTipoDescuentoDef.EnuReteFuente
                    lstrTipoImp = Format(CType(EnuTipoImpuestoMF.enuRetefte, Integer), "00")
                Case EnuTipoDescuentoDef.EnuReteIca
                    lstrTipoImp = Format(CType(EnuTipoImpuestoMF.enuIca, Integer), "00")
                Case EnuTipoDescuentoDef.EnuReteIva
                    lstrTipoImp = Format(CType(EnuTipoImpuestoMF.enuIva, Integer), "00")
                Case Else
                    Throw New ErrorInesperadoPanLException("Tipo de retención no esperado")
            End Select
        End If
        Return lstrTipoImp
    End Function
    Private Enum EnuTipoImpuestoMF As Integer
        None = 0
        enuIva
        enuImpoconsumo
        enuIca
        enuRetefte
    End Enum
    Friend Sub SPrepareParaReprocesarEfac()
        If GobjParametros.ObjAutorizaEFacBln.ObjValorPro AndAlso ObjIdEstadoEDocEnt.ObjValorPro <>
                EnuEstadoEDoc.EnuNoReg AndAlso ObjIdEstadoEDocEnt.ObjValorPro <>
                EnuEstadoEDoc.EnuNoEDoc AndAlso ObjIdEstadoEDocEnt.ObjValorPro <>
                EnuEstadoEDoc.EnuInvalida Then
            If Not String.IsNullOrEmpty(ObjCUDocStr.ToString) Then
                EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuErrorFtp
                ObjCUDEStr.ObjValorPro = ""
                ObjCUDocStr.ObjValorPro = ""
                SActualice(True)
            End If
        End If
    End Sub
#End Region
#Region "DataTable Novedades"
    Friend ReadOnly Property DtbNovedades As DataTable
        Get
            SCargueDtbNovedades()
            SComplementeTablaNov()
            Return MdtbNovedades
        End Get
    End Property
    Private Sub SCargueDtbNovedades()
        If IsNothing(MdtbNovedades) Then
            Dim lstrIdNotaCr = ObjIdNotaCrEnt.ToString
            If String.IsNullOrEmpty(lstrIdNotaCr) Then lstrIdNotaCr = "0"
            Dim lstrIndice = {{ClsPrefijoFact_NovStr.SstrNombreCampoBd, "ASC"},
                              {ClsIdFactura_NovEnt.SstrNombreCampoBd, "ASC"},
                              {ClsIdNovedadShr.SstrNombreCampoBd, "ASC"}}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsIdTipoDocOrigenByt.SstrNombreCampoBd &
                    " = " & EnuTipoDocOri.EnuNotaCr & " AND " &
                    ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd & " = '" & ObjPrefijo_NotaCrStr.ObjValorPro &
                    "' AND " & ClsIdDocOrigenEnt.SstrNombreCampoBd & " = " & lstrIdNotaCr
            Dim lstrCamposSelect() = {"*", "'' AS ConceptoNovedad", "'' AS NroFac"}
            MdtbNovedades = ClsPanorama.FdtbDataTable(ClsNovedad.SstrNombreTabla, lstrCamposSelect, lstrIndice,
                    lstrFiltro)
        End If
    End Sub
    Private Sub SComplementeTablaNov()
        Dim ldrwNovedades = MdtbNovedades.Select
        Dim lstrConceptoNovedad = String.Empty
        For Each ldrwNovedad As DataRow In ldrwNovedades
            Dim lstrPrefFac As String = ClsPanorama.FobjValorCampo(ldrwNovedad(ClsPrefijoFact_NovStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString)
            Dim lentIdFac As Integer = ClsPanorama.FobjValorCampo(ldrwNovedad(ClsIdFactura_NovEnt.SstrNombreCampoBd),
                    EnuTipoValor.enuInteger)
            Dim lstrNroFac = ClsPanorama.FstrNumeroDcto(lstrPrefFac, lentIdFac)
            Dim lenuTipoNovedad As EnuTipoNov =
                    ClsPanorama.FobjValorCampo(ldrwNovedad(ClsIdTipoNovedadByt.SstrNombreCampoBd),
                    EnuTipoValor.enuByte)
            Select Case lenuTipoNovedad
                Case EnuTipoNov.EnuDbCap
                    If ObjAnuladoBln.ObjValorPro Then
                        lstrConceptoNovedad = "Reversión del Débito. Nota Crédito Anulada"
                    End If
                Case EnuTipoNov.EnuDbIva
                    If ObjAnuladoBln.ObjValorPro Then
                        lstrConceptoNovedad = "Reversión del Iva. Nota Crédito Anulada"
                    End If
                Case EnuTipoNov.EnuCrDctoCap
                    lstrConceptoNovedad = "Devolución de Capital"
                Case EnuTipoNov.EnuCrDctoInt
                    lstrConceptoNovedad = "Devolución de Intereses de Mora"
                Case EnuTipoNov.EnuCrRetCre
                    lstrConceptoNovedad = "Retención del CREE"
                Case EnuTipoNov.EnuCrRetFte
                    lstrConceptoNovedad = "Retención en la fuente"
                Case EnuTipoNov.EnuCrRetIca
                    lstrConceptoNovedad = "Retención de Ind. y Comercio"
                Case EnuTipoNov.EnuCrRetIva
                    lstrConceptoNovedad = "Retención del IVA"
                Case EnuTipoNov.EnuRDbCap
                    lstrConceptoNovedad = "Venta Servicio Reversada"
                Case EnuTipoNov.EnuRDbIva
                    lstrConceptoNovedad = "IVA Generado Reversado"
                Case EnuTipoNov.EnuRDbIvaInt
                    lstrConceptoNovedad = "IVA Generado a Int. Mora Reversado"
                Case EnuTipoNov.EnuRDbInt
                    lstrConceptoNovedad = "Interes de Mora Reversado"
                Case EnuTipoNov.EnuRCrDctoCap
                    lstrConceptoNovedad = "Descuento Capital Reversado"
                Case EnuTipoNov.EnuRCrIvaGas
                    lstrConceptoNovedad = "Iva llevado al Gasto Reversado"
                Case EnuTipoNov.EnuRCrDctoInt
                    lstrConceptoNovedad = "Descuento Int. Mora Reversado"
                Case EnuTipoNov.EnuRCrRetFte
                    lstrConceptoNovedad = "Retefuente Reversada"
                Case EnuTipoNov.EnuRCrRetIva
                    lstrConceptoNovedad = "ReteIva Reversada"
                Case EnuTipoNov.EnuRCrRetIca
                    lstrConceptoNovedad = "ReteIca Reversada"
                Case EnuTipoNov.EnuRCrRetCre
                    lstrConceptoNovedad = "ReteCree Reversada"
                Case EnuTipoNov.EnuCrIvaGas
                    lstrConceptoNovedad = "IVA llevado al Gasto"
                Case Else
                    Throw New ErrorInesperadoPanLException("Tipo de Novedad inconsistente!")
            End Select
            ldrwNovedad("NroFac") = lstrNroFac
            ldrwNovedad("ConceptoNovedad") = lstrConceptoNovedad
        Next
    End Sub
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsComentario_NotaCrStr
    Inherits ClsCBPropiedad
    Private ReadOnly MobjPadre As ClsNotaCr = Nothing
    Private Const MCSTRNOMBRECAMPOBD As String = "Comentario"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Comentario"
        HshrLongitud = 500
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 3, ShrLongitud, BlnEsRequerido)
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            If Not HblnEsValido Then
                HstrMens = "Es necesario ingresar un Comentario!"
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
        If Not IsNothing(HobjValorPro) Then
            Return HobjValorPro
        Else
            Return ""
        End If
    End Function
End Class
Friend Class ClsFecha_NotaCrDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaNota"
    Private ReadOnly MobjPadre As ClsNotaCr = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "FechaNotaCr"
        HenuTipoValor = EnuTipoValor.enuDate
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MyBase.SVaciePropiedad()
        HobjValorNew = Date.Today
        HobjValorPro = Date.Today
    End Sub
    Friend Sub SValideFechaConItems()
        HblnEsValido = MobjPadre.FblnEsValidaFecha(HobjValorPro)
    End Sub
    Public Overrides Sub SValide()
        Dim ldtmFechaMin = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
        Dim ldtmFechaMax = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
        If ldtmFechaMax > Date.Today Then
            ldtmFechaMax = Now
        End If
        If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax,
                    BlnEsRequerido)
            HstrMens = String.Empty
            If Not HblnEsValido Then
                If HobjValorNew > Date.Today Then
                    HstrMens = "La fecha es posterior al día de Hoy!"
                Else
                    HstrMens = "La Fecha está por fuera del Período Actual!"
                End If
            End If
            If Not String.IsNullOrEmpty(HstrMens) Then
                SNotifiqueDatInv()
            End If
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
Friend Class ClsFechaAnulacion_NotaCrDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaAnulacion"
    Private ReadOnly MobjPadre As ClsNotaCr = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "FechaAnulacion_NotaCr"
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
        If HblnEsValido Then
            If HobjValorNew <> GCDTMFECHANULA Then
                HblnEsValido = MobjPadre.ObjAnuladoBln.ObjValorPro
                If HblnEsValido Then
                    Dim ldtmFechaMin = GCDTMFECHANULA
                    Dim ldtmFechaMax = GCDTMFECHANULA
                    If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                        HblnEsValido = MobjPadre.ObjAnuladoBln.ObjValorPro
                        If HblnEsValido Then
                            If Today > GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo Then
                                ldtmFechaMin = Today.AddDays(-Today.Day)
                                ldtmFechaMax = Date.Today
                            Else
                                ldtmFechaMin = Date.Today
                                ldtmFechaMax = Now
                            End If
                        End If
                    ElseIf MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                        ldtmFechaMin = HobjValorOriginal
                        ldtmFechaMax = HobjValorOriginal
                    End If
                    HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax,
                            BlnEsRequerido)
                End If
            End If
        End If
    End Sub
    Private Sub EPosSetValor() Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando AndAlso BlnEsValido Then
            If HobjValorPro <> GCDTMFECHANULA Then
                MobjPadre.ObjValor_NotaCrDec.SValide()
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
Friend Class ClsIdCliente_NotaCrDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTerceroCliente"
    Private MobjCliente As ClsCliente = Nothing
    Private ReadOnly MobjPadre As ClsNotaCr = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdTerceroCliente_NotaCr"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCDBLMINTERC, GCDBLMAXTERC,
                BlnEsRequerido)
        Dim lobjValorIng = HobjValorNew
        HstrMens = String.Empty
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                If IsNothing(MobjCliente) Then
                    MobjCliente = New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
                End If
                If HobjValorNew <> HobjValorPro OrElse Not MobjCliente.BlnExiste Then
                    Dim lobjLlavePrincipal() = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
                    MobjCliente.SAbra(lobjLlavePrincipal)
                End If
                MobjPadre.ObjClienteNotaCr = MobjCliente
                If Not MobjCliente.BlnExiste Then
                    HblnEsValido = False
                    MobjCliente.SVacie()
                    HstrMens = "La Id. del Cliente ingresada, '" &
                        lobjValorIng.ToString & "',  no es valida!"
                End If
            End If
        Else
            HstrMens = "La Id. del Cliente ingresada, '" &
                        lobjValorIng.ToString & "',  no es válida!"
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
Friend Class ClsIdNotaCrEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdNotaCr"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdNotaCredito"
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
Friend Class ClsIdPredioAgrupador_NotaCrStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdPredioAgrupador"
    Private ReadOnly MobjPadre As ClsNotaCr = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdPredioAgrupador NotaCr"
        HenuTipoValor = EnuTipoValor.enuString
        HshrLongitud = 20
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MyBase.SVaciePropiedad()
        HobjValorNew = "***"
        HobjValorPro = "***"
        HobjValorOriginal = "***"
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud,
                BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                HblnEsValido = HobjValorNew <> My.Resources.Ninguno AndAlso HobjValorNew <> "***"
                If HblnEsValido Then
                    If Not String.IsNullOrEmpty(HobjValorNew) Then
                        Dim lobjLlavePrincipal() As Object = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
                        Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
                        lobjPredio.SAbra(lobjLlavePrincipal)
                        HblnEsValido = lobjPredio.BlnExiste
                    End If
                End If
            Else
                HblnEsValido = HobjValorNew = HobjValorOriginal
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
            Return ObjValorPro.ToString
        Else
            Return ""
        End If
    End Function
End Class
Friend Class ClsIdTipoNotaCrByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTipoNotaCr"
    Private ReadOnly MobjPadre As ClsNotaCr = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "Tipo Nota Cr."
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuTipoNotaCrDef.EnuDescuento,
                EnuTipoNotaCrDef.EnuRetenciones, BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
                If HobjValorNew <> EnuTipoNotaCrDef.EnuDescuento Then
                    HblnEsValido = MobjPadre.ObjModoNotaCrByt.ObjValorPro =
                            EnuModoNotaCr.EnuPorFactura
                    If Not HblnEsValido Then
                        HstrMens = "El Modo General solo es permitido para notas crèdito por " &
                                "descuento!"
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
        If IsNothing(HobjValorPro) Then
            Return ""
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsModoNotaCrByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdModoNotaCr"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Modo Nota Cr."
        HenuTipoValor = EnuTipoValor.EnuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuModoNotaCr.EnuPorFactura,
                EnuModoNotaCr.EnuPorValor, BlnEsRequerido)
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
Friend Class ClsPrefijo_NotaCrStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PrefijoNotaCr"
    Private ReadOnly MobjPadre As ClsCBObjetoPan = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "PrefijoNotaCr"
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
                HblnEsValido = (HobjValorNew = GobjParametros.FstrPrefijoDoc(EnuTipoDocOri.EnuNotaCr))
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
            Return String.Empty
        Else
            Return HobjValorPro.ToString
        End If
    End Function
End Class
Friend Class ClsValor_NotaCrDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Valor"
    Private ReadOnly MobjPadre As ClsNotaCr = Nothing
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
Imports System.Drawing
Imports System.Windows.Media.Imaging
Imports ThoughtWorks.QRCode.Codec
Friend Class ClsNotaCon
#Region "Definiciones"
    Inherits ClsCBObjetoPan
    ' Constantes
    Private Const MCSTRNOMBRETABLA As String = "OriNotasCon"
    ' Variables de modulo
    Private McolItemsNotaCon As Collection = Nothing
    Private MdtbItemsNotaCon As DataTable = Nothing
    Private MdtbNovedadesNCon As DataTable = Nothing
    Private McolNovedades As Collection = Nothing
    Private MobjClienteNotaCon As ClsCliente = Nothing
    Private MobjPredioAgrNCon As ClsPredio = Nothing
    Private MobjAnticipo As ClsAnticipo = Nothing
#End Region

#Region "Constructores"
    ''' <summary>
    ''' Instancia un objeto Nota Contable (Aplicación Anticipos) en modo único
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
    ''' Instancia un objeto Nota Contable (Aplicación Anticipos) en modo navegable
    ''' </summary>
    Public Sub New(astrPref As String)
        HobjPadre = Nothing
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion
        lstrFiltro &= " AND " & ClsPrefijo_NotaConStr.SstrNombreCampoBd & " = '" & astrPref & "'"
        HcolFiltros.Add(lstrFiltro)
        Dim lstrCamposSelect = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsPrefijo_NotaConStr.SstrNombreCampoBd, ClsIdNotaConEnt.SstrNombreCampoBd}
        HblnEsCreable = False
        HblnEsSuprimible = False
        HblnEsModificable = False
        HblnEsAnulable = True
        HcolTablas.Add(MCSTRNOMBRETABLA)
        HcolCamposSelect.Add(lstrCamposSelect)
    End Sub
    ''' <summary>
    ''' Instancia el objeto como un objeto no navegable, básicamente para formar parte de una colección
    ''' </summary>
    ''' <param name="aobjPadre">Objeto Administrador al cual pertenece el objeto que se esta instanciando</param>
    ''' <param name="adrwNotaCon">DataRow que contiene los valores de las propiedades del objeto</param>
    ''' <remarks></remarks>
    Public Sub New(aobjPadre As ClsCliente, adrwNotaCon As DataRow)
        HobjPadre = aobjPadre
        HenuTipoObjeto = EnuModoInstanciaObjDef.enuDeColeccion
        HblnEsSuprimible = False
        '
        DrwRegistroActual = adrwNotaCon
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
            Return EnuIdClasesPanDef.EnuNotaCon
        End Get
    End Property
    Protected Overrides ReadOnly Property HstrNombreClase As String
        Get
            Return "Nota Contable"
        End Get
    End Property
#End Region
#Region "Propiedades Prop"
    Friend ReadOnly Property ObjCUDEStr As New ClsCUDEStr(Me)
    Friend ReadOnly Property ObjCUDocStr As New ClsCUDocStr(Me)
    Friend ReadOnly Property ObjFechaAnulacion_NotaConDtm As New ClsFechaAnulacion_NotaConDtm(Me)
    Friend ReadOnly Property ObjFecha_NotaConDtm As New ClsFecha_NotaConDtm(Me)
    Friend ReadOnly Property ObjIdAnticipo_NotaConEnt As New ClsIdAnticipo_NotaConEnt(Me)
    Friend ReadOnly Property ObjIdCarpeta_NotaConShr As New ClsIdCarpetaShr(Me)
    Friend ReadOnly Property ObjIdCentroUtil_NotaConShr As New ClsIdCentroUtilShr(Me)
    Friend ReadOnly Property ObjIdCliente_NotaConDbl As New ClsIdCliente_NotaConDbl(Me)
    Friend ReadOnly Property ObjIdEstadoEDocEnt As New ClsIdEstadoEDocEnt(Me)
    Friend ReadOnly Property ObjIdNotaConEnt As New ClsIdNotaConEnt(Me)
    Friend ReadOnly Property ObjIdNotaRCrEnt As New ClsIdNotaRCrEnt(Me)
    Friend ReadOnly Property ObjIdPredioAgrupador_NotaConStr As New ClsIdPredioAgrupador_NotaConStr(Me)
    Friend ReadOnly Property ObjIdUsuario_NotaConStr As New ClsIdUsuarioStr(Me)
    Friend ReadOnly Property ObjPrefijo_NotaConStr As New ClsPrefijo_NotaConStr(Me)
    Friend ReadOnly Property ObjPrefijo_NotaRCrStr As New ClsPrefijo_NotaRCrStr(Me)
    Friend ReadOnly Property ObjValor_NotaConDec As New ClsValor_NotaConDec(Me)
    Friend Overrides ReadOnly Property ColPropiedades As Collection
        Get
            If HcolPropiedades.Count = 0 Then
                HcolPropiedades.Add(ObjAnuladoBln)
                HcolPropiedades.Add(ObjFechaCreacionDtm)
                HcolPropiedades.Add(ObjIdUsuarioAnuloStr)
                HcolPropiedades.Add(ObjOrigenInstanciaStr)
                HcolPropiedades.Add(ObjOrigenInstanciaAnuloStr)
                HcolPropiedades.Add(ObjCUDEStr)
                HcolPropiedades.Add(ObjCUDocStr)
                HcolPropiedades.Add(ObjFechaAnulacion_NotaConDtm)
                HcolPropiedades.Add(ObjFecha_NotaConDtm)
                HcolPropiedades.Add(ObjIdAnticipo_NotaConEnt)
                HcolPropiedades.Add(ObjIdCarpeta_NotaConShr)
                HcolPropiedades.Add(ObjIdCentroUtil_NotaConShr)
                HcolPropiedades.Add(ObjIdCliente_NotaConDbl)
                HcolPropiedades.Add(ObjIdEstadoEDocEnt)
                HcolPropiedades.Add(ObjIdNotaConEnt)
                HcolPropiedades.Add(ObjIdNotaRCrEnt)
                HcolPropiedades.Add(ObjIdPredioAgrupador_NotaConStr)
                HcolPropiedades.Add(ObjIdUsuario_NotaConStr)
                HcolPropiedades.Add(ObjPrefijo_NotaConStr)
                HcolPropiedades.Add(ObjPrefijo_NotaRCrStr)
                HcolPropiedades.Add(ObjValor_NotaConDec)
            End If
            Return HcolPropiedades
        End Get
    End Property
#End Region
#Region "Otras propiedades"
    ''' <summary>
    ''' Devuelve uns string compuesto por el prefijo de la nota y el id de la nota separados por un
    ''' guion. Si no existe el prefijo devuelve solo el id de la factura
    ''' </summary>
    ''' <value></value>
    Friend ReadOnly Property StrNumeroNotaCon As String
        Get
            Dim lstrNumeroNotaCon As String = ClsPanorama.FstrNumeroDcto(ObjPrefijo_NotaConStr.ObjValorPro,
                    ObjIdNotaConEnt.ObjValorPro)
            Return lstrNumeroNotaCon
        End Get
    End Property

    Friend ReadOnly Property StrNumeroNotaRCr As String
        Get
            Dim lstrNumeroNotaRCr As String = ClsPanorama.FstrNumeroDcto(ObjPrefijo_NotaRCrStr.ObjValorPro,
                    ObjIdNotaRCrEnt.ObjValorPro)
            Return lstrNumeroNotaRCr
        End Get
    End Property

    Friend ReadOnly Property ObjClienteNotaCon As ClsCliente
        Get
            If IsNothing(MobjClienteNotaCon) Then
                MobjClienteNotaCon = New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
                MobjClienteNotaCon.SAbra({GshrIdCarpeta, GshrIdCentroUtil, ObjIdCliente_NotaConDbl.ObjValorPro})
            End If
            Return MobjClienteNotaCon
        End Get
    End Property

    Friend ReadOnly Property ObjPredioAgrNCon As ClsPredio
        Get
            If IsNothing(MobjPredioAgrNCon) Then
                If Not String.IsNullOrEmpty(ObjIdPredioAgrupador_NotaConStr.ToString) AndAlso
                        ObjIdPredioAgrupador_NotaConStr.ToString <> GCSTRSINPA Then
                    MobjPredioAgrNCon = New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
                    Dim lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil,
                            ObjIdPredioAgrupador_NotaConStr.ObjValorPro}
                    MobjPredioAgrNCon.SAbra(lobjValorLlave)
                    If Not MobjPredioAgrNCon.BlnExiste Then
                        MobjPredioAgrNCon = Nothing
                    End If
                End If
            End If
            Return MobjPredioAgrNCon
        End Get
    End Property

    Friend ReadOnly Property BlnEsAjusteCuotaAdmin As Boolean
        Get
            If IsNothing(MobjAnticipo) Then
                MobjAnticipo = New ClsAnticipo(EnuModoInstanciaObjDef.enuUnico)
                Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil,
                        ObjIdAnticipo_NotaConEnt.ObjValorPro}
                MobjAnticipo.SAbra(lobjValorLlave)
            End If
            Return MobjAnticipo.BlnEsAnticipoAjuste
        End Get
    End Property

    Friend ReadOnly Property ObjFacturaAfectada As ClsFactura
        Get
            If BlnEsAjusteCuotaAdmin() Then
                Dim lobjItNotCon As ClsItemNotaCon = ColItemsNotaCon(1)
                Dim lstrPrefFac As String = lobjItNotCon.ObjPrefijoFact_ItemNotaConStr.ToString
                Dim lentIdFac As Integer = lobjItNotCon.ObjIdFactura_ItemNotaConEnt.ObjValorPro
                Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, lstrPrefFac, lentIdFac}
                Dim lobjFac As New ClsFactura()
                lobjFac.SAbra(lobjValorLlave)
                Return lobjFac
            Else
                Return Nothing
            End If
        End Get
    End Property
#End Region
#End Region

#Region "Procedimientos y funciones invalidantes"
    Protected Overrides Sub SVacie()
        MyBase.SVacie()
        MobjClienteNotaCon = Nothing
        MobjPredioAgrNCon = Nothing
        MdtbItemsNotaCon = Nothing
        McolItemsNotaCon = Nothing
        MdtbNovedadesNCon = Nothing
        McolNovedades = Nothing
        MobjAnticipo = Nothing
    End Sub

    Protected Overrides Sub SInicialiceObj()
        ObjAnuladoBln.ObjValorPro = False
        ObjFechaAnulacion_NotaConDtm.ObjValorPro = GCDTMFECHANULA
        ObjFechaCreacionDtm.ObjValorPro = Date.Now
        ObjIdAnticipo_NotaConEnt.ObjValorPro = 0
        ObjIdCarpeta_NotaConShr.ObjValorPro = GshrIdCarpeta
        ObjIdCentroUtil_NotaConShr.ObjValorPro = GshrIdCentroUtil
        ObjIdUsuario_NotaConStr.ObjValorPro = GstrIdUsuario
        ObjIdUsuarioAnuloStr.ObjValorPro = String.Empty
        ObjOrigenInstanciaStr.ObjValorPro = GstrOrigenActual
        ObjOrigenInstanciaAnuloStr.ObjValorPro = String.Empty
        ObjPrefijo_NotaConStr.ObjValorPro = GobjParametros.FstrPrefijoDoc(
                    EnuTipoDocOri.EnuNotaCon)
        ObjValor_NotaConDec.ObjValorPro = 0
        ObjCUDocStr.ObjValorPro = String.Empty
        ObjCUDEStr.ObjValorPro = String.Empty
        ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuNoEDoc
    End Sub

    Public Overrides Function FblnEsAnulable() As Boolean
        Dim lblnEsAnulable = BlnEsAnulable AndAlso BlnExiste
        If lblnEsAnulable Then
            If GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro Then
                lblnEsAnulable = Date.Today <=
                        GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
            End If
        End If
        If lblnEsAnulable Then
            Dim lstrPeriodoNCon = ClsPanorama.FstrPeriodo(ObjFecha_NotaConDtm.ObjValorPro)
            Dim lstrPeriodoActual = GobjParametros.ObjAnoActual.StrIdPeriodoActual
            lblnEsAnulable = lstrPeriodoNCon = lstrPeriodoActual
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
                ObjFechaAnulacion_NotaConDtm.ObjValorPro = Now
            Else
                Dim ldtmFecFinPer As Date =
                        GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
                If ldtmFecFinPer < Date.Today Then
                    ObjFechaAnulacion_NotaConDtm.ObjValorPro = ldtmFecFinPer
                Else
                    ObjFechaAnulacion_NotaConDtm.ObjValorPro = Now
                End If
            End If
            ' Items nota, Factura, items factura
            Dim lstrPrefFact As String, lentIdFact As Integer, lshrIdItemFac As Short
            Dim ldecVlrItem As Decimal
            For Each lobjItemNCo As ClsItemNotaCon In ColItemsNotaCon
                ldecVlrItem = lobjItemNCo.ObjValor_ItemNotaConDec.ObjValorPro
                lobjItemNCo.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                lobjItemNCo.ObjAnuladoBln.ObjValorPro = True
                lobjItemNCo.ObjValor_ItemNotaConDec.ObjValorPro = 0
                lstrPrefFact = lobjItemNCo.ObjPrefijoFact_ItemNotaConStr.ObjValorPro
                lentIdFact = lobjItemNCo.ObjIdFactura_ItemNotaConEnt.ObjValorPro
                lshrIdItemFac = lobjItemNCo.ObjIdItemFac_ItemNotaConShr.ObjValorPro
                SReverseValoresFact(lstrPrefFact, lentIdFact, lshrIdItemFac, ldecVlrItem)
            Next
            ' Anticipo
            Dim lentIdAnti = ObjIdAnticipo_NotaConEnt.ObjValorPro
            Dim lobjAnticipo As New ClsAnticipo(EnuModoInstanciaObjDef.enuUnico)
            Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, lentIdAnti}
            lobjAnticipo.SAbra(lobjValorLlave)
            lobjAnticipo.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
            lobjAnticipo.ObjDebitos_AntDec.ObjValorPro -= ObjValor_NotaConDec.ObjValorPro
            lobjAnticipo.SActualice(True)
            ' Novedades anticipo
            For Each lobjNovAnt As ClsNovedadAnticipo In lobjAnticipo.ColNovedadesAnt
                If lobjNovAnt.ObjIdTipoDocOrigen_NovAntByt.ObjValorPro =
                        EnuTipoDocOri.EnuNotaCon AndAlso
                        lobjNovAnt.ObjPrefijoDocOrigen_NovAntStr.ObjValorPro =
                        ObjPrefijo_NotaConStr.ObjValorPro AndAlso
                        lobjNovAnt.ObjIdDocOrigen_NovAntEnt.ObjValorPro =
                        ObjIdNotaConEnt.ObjValorPro Then
                    lobjNovAnt.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                    lobjNovAnt.ObjAnuladoBln.ObjValorPro = True
                    lobjNovAnt.ObjValor_NovAntDec.ObjValorPro = 0
                    lobjNovAnt.SActualice(True)
                End If
            Next
            'Novedades
            For Each lobjNov As ClsNovedad In ColNovedades
                lobjNov.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
                lobjNov.ObjAnuladoBln.ObjValorPro = True
                lobjNov.ObjValor_NovDec.ObjValorPro = 0
                lobjNov.SActualice(True)
            Next
            ObjValor_NotaConDec.ObjValorPro = 0
        End If
        Return lblnAnulado
    End Function

    Private Sub SReverseValoresFact(astrPrefFac As String, aentIdFac As Integer,
            ashrIdItemFac As Short, adecVlr As Decimal)
        Dim lobjFac As New ClsFactura()
        Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, astrPrefFac,
                aentIdFac}
        lobjFac.SAbra(lobjValorLlave)
        lobjFac.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
        lobjFac.ObjCreditos_FactDec.ObjValorPro -= adecVlr
        Dim lobjItemFac As ClsItemFactura = lobjFac.ColItemsFactura(
                ashrIdItemFac.ToString())
        lobjItemFac.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
        lobjItemFac.ObjCreditos_ItemFactDec.ObjValorPro -= adecVlr
        lobjFac.SActualice(True)
    End Sub

    Protected Overrides Sub SActualice(ablnExigeRequeridos As Boolean)
        Dim lblnNoHayError = False
        GobjPanDat.SControleProcesoObj(True)
        Try
            GobjPanDat.SInicialiceTransaccion()
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuCreando Then
                SNumereObj()
                SComplementeEFac()
            End If
            ClsPanorama.SActualiceCol(ColItemsNotaCon)
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

    Friend Overrides ReadOnly Property StrIdObjeto As String
        Get
            Return ClsPanorama.FstrNumeroDcto(ObjPrefijo_NotaConStr.ToString(),
                    ObjIdNotaConEnt.ObjValorPro)
        End Get
    End Property
#End Region

#Region "Procedimientos del objeto"
    Private Sub SNumereObj()
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
            Dim lentIdNotaCon As Integer
            Dim lstrPrefijo = GobjParametros.FstrPrefijoDoc(EnuTipoDocOri.EnuNotaCon)
            If IsNothing(lstrPrefijo) Then lstrPrefijo = String.Empty
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                    ClsPrefijo_NotaConStr.SstrNombreCampoBd & " = '" & lstrPrefijo & "'"
            lentIdNotaCon = ClsPanorama.FobjUltimaIdNumericaObjeto(SstrNombreTabla,
                    ClsIdNotaConEnt.SstrNombreCampoBd, ObjIdNotaConEnt.EnuTipoValor,
                    lstrFiltro)
            If lentIdNotaCon < GobjParametros.FentNumeracionInicialDoc(EnuTipoDocOri.EnuNotaCon) Then
                lentIdNotaCon = GobjParametros.FentNumeracionInicialDoc(EnuTipoDocOri.EnuNotaCon)
            End If
            lentIdNotaCon += 1
            ObjPrefijo_NotaConStr.ObjValorPro = lstrPrefijo
            ObjIdNotaConEnt.ObjValorPro = lentIdNotaCon
            For Each lobjItemNotaCon As ClsItemNotaCon In ColItemsNotaCon
                lobjItemNotaCon.ObjIdNotaCon_ItemNotaConEnt.ObjValorPro = lentIdNotaCon
                lobjItemNotaCon.ObjPrefijoNotaCon_ItemNotaConStr.ObjValorPro = lstrPrefijo
            Next
        End If
    End Sub

    Private Sub SComplementeEFac()
        If GobjParametros.ObjAutorizaEFacBln.ObjValorPro Then
            If FblnInsertarEFac() Then
                ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuNoReg
            End If
        End If
    End Sub

    Friend Sub SReverse(astrPrefNotaRCr As String, aentIdNotaRCr As Integer,
                        adtmFechaRRC As Date)
        If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
            HblnEsAnulable = True
            EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
        End If
        ObjAnuladoBln.ObjValorPro = True
        ObjIdUsuarioAnuloStr.ObjValorPro = GstrIdUsuario
        ObjOrigenInstanciaAnuloStr.ObjValorPro = GstrOrigenActual
        ObjPrefijo_NotaRCrStr.ObjValorPro = astrPrefNotaRCr
        ObjIdNotaRCrEnt.ObjValorPro = aentIdNotaRCr
        ObjFechaAnulacion_NotaConDtm.ObjValorPro = adtmFechaRRC
        SReverseNovedades(astrPrefNotaRCr, aentIdNotaRCr, adtmFechaRRC)
        SActualice(True)
    End Sub

    Private Sub SReverseNovedades(astrPrefNotaRCr As String, aentIdNotaRCr As Integer,
                                  adtmFechaNotaRCr As Date)
        Dim lentIdFac As Integer, lstrPrefFac As String
        Dim ldecValorNov As Decimal
        ' El único documento que puede reversar una nota contable( aplicacion anticipos) es la Nota Reversión RC
        Dim lenuTipoDocOrigen = EnuTipoDocOri.EnuNotaRevCr
        Dim lobjFactura As New ClsFactura()
        For Each lobjNov As ClsNovedad In ColNovedades
            ldecValorNov = lobjNov.ObjValor_NovDec.ObjValorPro
            lstrPrefFac = lobjNov.ObjPrefijoFact_NovStr.ObjValorPro
            lentIdFac = lobjNov.ObjIdFactura_NovEnt.ObjValorPro
            lobjFactura.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefFac, lentIdFac})
            lobjFactura.SReverseNovedad(lobjNov, lenuTipoDocOrigen, adtmFechaNotaRCr,
                        astrPrefNotaRCr, aentIdNotaRCr)
            lobjFactura.SActualice(True)
        Next
    End Sub

    Friend Function FdecVlrAntApliItemFac(astrIdItemFac As String) As Decimal
        Dim ldecDscto = 0D
        For Each lobjNov As ClsNovedad In ColNovedades
            If lobjNov.ObjIdItemFact_NovShr.ToString = astrIdItemFac Then
                If lobjNov.ObjIdTipoNovedadByt.ObjValorPro = EnuTipoNov.EnuCrAnApCap Then
                    ldecDscto += lobjNov.ObjValor_NovDec.ObjValorPro
                End If
            End If
        Next
        Return ldecDscto
    End Function

    Friend Function FarlCorreosNCon() As ArrayList
        Dim larlListaCorreos As New ArrayList
        Dim lstrCorreoCli As String, lstrCorreoPredio As String
        If ObjClienteNotaCon.ObjRecibeDocsPorEmailBln.ObjValorPro Then
            lstrCorreoCli = ObjClienteNotaCon.ObjEmailStr.ToString
            If Not String.IsNullOrEmpty(lstrCorreoCli) Then
                larlListaCorreos.Add(lstrCorreoCli)
            End If
        End If
        If Not IsNothing(ObjPredioAgrNCon) Then
            lstrCorreoPredio = ObjPredioAgrNCon.ObjEmailAdiStr.ToString
            If Not String.IsNullOrEmpty(lstrCorreoPredio) Then
                larlListaCorreos.Add(lstrCorreoPredio)
            End If
        End If
        Return larlListaCorreos
    End Function

    Friend Function FbytQRNcon() As Byte()
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
        Dim ldtmFecNCon As Date = ObjFecha_NotaConDtm.ObjValorPro
        Dim lstrFecNCon As String = Year(ldtmFecNCon).ToString &
                Format(Month(ldtmFecNCon), "00") & Format(Day(ldtmFecNCon), "00") &
                Format(Hour(ldtmFecNCon), "00") & Format(Minute(ldtmFecNCon), "00") &
                Format(Second(ldtmFecNCon), "00")
        Dim lstrNCon As String = "Número:" & StrNumeroNotaCon & vbCrLf &
                "Fecha:" & lstrFecNCon & vbCrLf &
                "Nit:" & lstrNit & vbCrLf &
                "DocAdq:" & ObjIdCliente_NotaConDbl.ToString & vbCrLf &
                "ValNcr:" & Format(ObjValor_NotaConDec.ObjValorPro, "#0.00") & vbCrLf &
                "ValIva:" & Format(0, "#0.00") & vbCrLf &
                "ValOtroIm:" & "0.00" & vbCrLf &
                "ValNcrIm:" & Format(ObjValor_NotaConDec.ObjValorPro, "#0.00") & vbCrLf
        If GobjParametros.ObjIdProveedorEFacEnt.ObjValorPro > EnuProveedorEFac.None Then
            lstrNCon &= "CUDE:" & ObjCUDEStr.ObjValorPro
        End If
        Try
            Dim lbmiQRNdb As New BitmapImage
            lbmiQRNdb.BeginInit()
            Dim lbtmQR As Bitmap = lqreQRNdb.Encode(lstrNCon, System.Text.Encoding.UTF8)
            lbytQR = FBytQR(lbtmQR)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly)
        End Try
        Return lbytQR
    End Function
#End Region

#Region "Manejo Items Nota Contable"
    Friend ReadOnly Property ColItemsNotaCon As Collection
        Get
            If IsNothing(McolItemsNotaCon) Then
                McolItemsNotaCon = New Collection
                If ObjIdNotaConEnt.BlnEsValido AndAlso EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuCreando Then
                    SCargueDtbItemsNotaCon()
                    If Not IsNothing(MdtbItemsNotaCon) AndAlso MdtbItemsNotaCon.Rows.Count > 0 Then
                        Dim ldrwItemsNotaCon() As DataRow = MdtbItemsNotaCon.Select
                        For Each ldrwItemNotCon As DataRow In ldrwItemsNotaCon
                            Dim lobjItemNotaCon As New ClsItemNotaCon(Me, ldrwItemNotCon)
                            lobjItemNotaCon.SLeaValores(True)
                            McolItemsNotaCon.Add(lobjItemNotaCon, lobjItemNotaCon.ObjIdItemNotaConShr.ToString)
                        Next
                    End If
                End If
            End If
            Return McolItemsNotaCon
        End Get
    End Property

    Friend ReadOnly Property DtbItemsNotaCon As DataTable
        Get
            SCargueDtbItemsNotaCon()
            SComplementeTablaItems()
            Return MdtbItemsNotaCon
        End Get
    End Property

    Private Sub SCargueDtbItemsNotaCon()
        If IsNothing(MdtbItemsNotaCon) Then
            Dim lstrIndice = {{ClsPrefijo_NotaConStr.SstrNombreCampoBd, "ASC"},
                              {ClsIdNotaCon_ItemNotaConEnt.SstrNombreCampoBd, "ASC"},
                              {ClsIdItemNotaConShr.SstrNombreCampoBd, "ASC"}}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsPrefijo_NotaConStr.SstrNombreCampoBd &
                    " = '" & ObjPrefijo_NotaConStr.ObjValorPro & "' AND " &
                    ClsIdNotaCon_ItemNotaConEnt.SstrNombreCampoBd & " = "
            If EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                lstrFiltro &= "0"
            Else
                If Not IsNothing(ObjIdNotaConEnt.ObjValorPro) Then
                    lstrFiltro &= ObjIdNotaConEnt.ObjValorPro
                Else
                    lstrFiltro &= "0"
                End If
            End If
            Dim lstrCamposSelect() = {"*", "'' AS NroFact", "'' AS Detalle"}
            MdtbItemsNotaCon = ClsPanorama.FdtbDataTable(ClsItemNotaCon.SstrNombreTabla, lstrCamposSelect,
                    lstrIndice, lstrFiltro)
        End If
    End Sub

    Friend Sub SGenereItemNotaCon(adecValor As Decimal, astrPrefijoFact As String,
            aentIdFactura As Integer, aenuTipoItemNotaCon As EnuTipoItemNotaConDef,
            ashrIdItemFac As Short)
        Dim lobjItemNotaCon As ClsItemNotaCon = FobjNuevoItemNotaCon()
        With lobjItemNotaCon
            .ObjPrefijoFact_ItemNotaConStr.ObjValorPro = astrPrefijoFact
            .ObjIdFactura_ItemNotaConEnt.ObjValorPro = aentIdFactura
            .ObjIdItemFac_ItemNotaConShr.ObjValorPro = ashrIdItemFac
            .ObjValor_ItemNotaConDec.ObjValorPro = adecValor
            .ObjIdTipoItemNotaConByt.ObjValorPro = aenuTipoItemNotaCon
        End With
        ObjValor_NotaConDec.ObjValorPro += adecValor
        McolItemsNotaCon.Add(lobjItemNotaCon, lobjItemNotaCon.ObjIdItemNotaConShr.ObjValorPro)
    End Sub

    Private Function FobjNuevoItemNotaCon() As ClsItemNotaCon
        Dim lobjItemNotaCon As ClsItemNotaCon = Nothing
        McolItemsNotaCon = ColItemsNotaCon
        SCargueDtbItemsNotaCon()
        Dim ldrwNewItem As DataRow = MdtbItemsNotaCon.NewRow
        lobjItemNotaCon = New ClsItemNotaCon(Me, ldrwNewItem)
        Dim lblnModificoPermisos = False
        With lobjItemNotaCon
            If Not CType(.EnuPermisosObj And EnuPermisosDef.EnuCrear, Boolean) Then
                .EnuPermisosObj += EnuPermisosDef.EnuCrear
                lblnModificoPermisos = True
            End If
            .SCreeObj(Nothing)
            .ObjAnuladoBln.ObjValorPro = False
            .ObjFechaCreacionDtm.ObjValorPro = Date.Now
            .ObjIdCarpeta_ItemNotaConShr.ObjValorPro = GshrIdCarpeta
            .ObjIdCentroUtil_ItemNotaConShr.ObjValorPro = GshrIdCentroUtil
            .ObjIdFactura_ItemNotaConEnt.ObjValorPro = 0
            .ObjIdItemNotaConShr.ObjValorPro = McolItemsNotaCon.Count + 1
            .ObjIdNotaCon_ItemNotaConEnt.ObjValorPro = 0
            .ObjPrefijoFact_ItemNotaConStr.ObjValorPro = String.Empty
            .ObjPrefijoNotaCon_ItemNotaConStr.ObjValorPro =
                    GobjParametros.FstrPrefijoDoc(EnuTipoDocOri.EnuNotaCon)
            .ObjValor_ItemNotaConDec.ObjValorPro = 0
            If lblnModificoPermisos Then
                .EnuPermisosObj -= EnuPermisosDef.EnuCrear
            End If
        End With
        Return lobjItemNotaCon
    End Function

    Private Sub SComplementeTablaItems()
        Dim ldrwItemsNotaCon As DataRow() = MdtbItemsNotaCon.Select()
        Dim lstrPrefFact As String, lentIdFact As Integer, lstrNroFact As String
        Dim lenuIdTipoNotaDb As EnuTipoItemNotaConDef
        For Each ldrwItNotaCon As DataRow In ldrwItemsNotaCon
            lstrPrefFact = ClsPanorama.FobjValorCampo(ldrwItNotaCon(ClsPrefijoFact_ItemNotaConStr.SstrNombreCampoBd),
                            EnuTipoValor.enuString)
            lentIdFact = ClsPanorama.FobjValorCampo(ldrwItNotaCon(ClsIdFactura_ItemNotaConEnt.SstrNombreCampoBd),
                            EnuTipoValor.enuInteger)
            lstrNroFact = ClsPanorama.FstrNumeroDcto(lstrPrefFact, lentIdFact)
            ldrwItNotaCon("NroFact") = lstrNroFact
            lenuIdTipoNotaDb = ClsPanorama.FobjValorCampo(ldrwItNotaCon(
                    ClsIdTipoItemNotaConByt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
            ldrwItNotaCon("Detalle") = FstrDetalle(lenuIdTipoNotaDb)
        Next
    End Sub

    Friend Shared Function FstrDetalle(aenuTipoItemNCon As EnuTipoItemNotaConDef) As String
        Dim lstrDetalle = String.Empty
        Select Case aenuTipoItemNCon
            Case EnuTipoItemNotaConDef.EnuAplicaAntCap
                lstrDetalle = "Anticipo aplicado a Capital"
            Case EnuTipoItemNotaConDef.EnuAplicaAntInt
                lstrDetalle = "Anticipo aplicado a Int de Mora"
            Case EnuTipoItemNotaConDef.EnuDsctoPP
                lstrDetalle = "Descuento por Pronto Pago aplicado"
            Case EnuTipoItemNotaConDef.EnuReteFuente
                lstrDetalle = "Retención en la Fuente aplicada"
            Case EnuTipoItemNotaConDef.EnuReteIca
                lstrDetalle = "Retención de Industria y Comercio aplicada"
            Case EnuTipoItemNotaConDef.EnuReteIva
                lstrDetalle = "Retención del Iva aplicada"
        End Select
        Return lstrDetalle
    End Function
    ''' <summary>
    ''' Devuelve el valor del anticipo aplicado por la nota sin tener en cuenta las retenciones y descuento por
    ''' pronto pago
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Friend Function FdecValorAntAplicado() As Decimal
        Dim ldecVlrAntAplicado = 0D
        For Each lobjItemNotaCon As ClsItemNotaCon In ColItemsNotaCon
            If lobjItemNotaCon.ObjIdTipoItemNotaConByt.ObjValorPro = EnuTipoItemNotaConDef.EnuAplicaAntCap OrElse
                    lobjItemNotaCon.ObjIdTipoItemNotaConByt.ObjValorPro =
                    EnuTipoItemNotaConDef.EnuAplicaAntInt Then
                ldecVlrAntAplicado += lobjItemNotaCon.ObjValor_ItemNotaConDec.ObjValorPro
            End If
        Next
        Return ldecVlrAntAplicado
    End Function
#End Region

#Region "eFac"
    ''' <summary>
    ''' Indica si es una factura electrónica y si esta registrada 
    ''' ''' </summary>
    Friend ReadOnly Property BlnEstaRegEFac As Boolean
        Get
            Return BlnEsDocEle AndAlso ObjIdEstadoEDocEnt.ObjValorPro >= EnuEstadoEDoc.EnuRegi
        End Get
    End Property

    Friend ReadOnly Property BlnEsDocEle As Boolean
        Get
            Dim lblnEsDocEle = GobjParametros.BlnEFacAutorizado
            If lblnEsDocEle Then
                lblnEsDocEle = BlnEsAjusteCuotaAdmin AndAlso ObjIdEstadoEDocEnt.ObjValorPro <>
                        EnuEstadoEDoc.EnuNoEDoc
            End If
            Return lblnEsDocEle
        End Get
    End Property

    Friend Function FblnInsertarEFac() As Boolean
        Dim lblnReg = False
        If BlnEsDocEle Then
            lblnReg = (ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuNoReg)
            If lblnReg Then
                For Each lobjItemNCon As ClsItemNotaCon In ColItemsNotaCon
                    lblnReg = lobjItemNCon.ObjFactura.BlnEsFacEle AndAlso
                            (Not lobjItemNCon.ObjFactura.FenuVerFacEFac = EnuVerEFac.EnuV1) AndAlso
                            lobjItemNCon.ObjFactura.ObjIdEstadoEDocEnt.ObjValorPro >=
                            EnuEstadoEDoc.EnuRegi AndAlso
                            lobjItemNCon.ObjFactura.ObjIdEstadoEDocEnt.ObjValorPro <>
                            EnuEstadoEDoc.EnuNoEDoc
                    If lblnReg Then Exit For
                Next
            End If
        End If
        Return lblnReg
    End Function

    Friend Function FblnActualizarEstEFac() As Boolean
        Dim lblnActu = False
        If BlnEsDocEle Then
            lblnActu = (ObjIdEstadoEDocEnt.ObjValorPro < EnuEstadoEDoc.EnuEnviada) AndAlso
                    (ObjIdEstadoEDocEnt.ObjValorPro >= EnuEstadoEDoc.EnuEnProceso)
        End If
        Return lblnActu
    End Function

    Friend Sub SHabiliteProcesarEFac()
        If BlnEsDocEle Then
            If ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuInvalida OrElse
                    ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuErrorFtp Then
                If EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
                    EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
                End If
                ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuNoReg
                SActualice(True)
            End If
        End If
    End Sub

    Friend Sub SPrepareParaReprocesarEfac()
        If GobjParametros.ObjAutorizaEFacBln.ObjValorPro AndAlso ObjIdEstadoEDocEnt.ObjValorPro <>
                EnuEstadoEDoc.EnuNoReg AndAlso ObjIdEstadoEDocEnt.ObjValorPro <>
                EnuEstadoEDoc.EnuNoEDoc AndAlso ObjIdEstadoEDocEnt.ObjValorPro <>
                EnuEstadoEDoc.EnuInvalida Then
            If Not String.IsNullOrEmpty(ObjCUDocStr.ToString) Then
                EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
                ObjIdEstadoEDocEnt.ObjValorPro = EnuEstadoEDoc.EnuErrorFtp
                ObjCUDEStr.ObjValorPro = ""
                ObjCUDocStr.ObjValorPro = ""
                SActualice(True)
            End If
        End If
    End Sub
#End Region

#Region "Novedades"
    Friend ReadOnly Property ColNovedades As Collection
        Get
            If IsNothing(McolNovedades) Then
                McolNovedades = New Collection
                SCargueDtbNovedadesNota()
                If Not IsNothing(MdtbNovedadesNCon) AndAlso MdtbNovedadesNCon.Rows.Count > 0 Then
                    For Each ldrwNovedad As DataRow In MdtbNovedadesNCon.Rows
                        Dim lobjNovedad As New ClsNovedad(Me, ldrwNovedad)
                        lobjNovedad.SLeaValores(True)
                        McolNovedades.Add(lobjNovedad)
                    Next
                End If
            End If
            Return McolNovedades
        End Get
    End Property

    Friend ReadOnly Property DtbNovedadesNotaCon As DataTable
        Get
            SCargueDtbNovedadesNota()
            SComplementeTablaNov()
            Return MdtbNovedadesNCon
        End Get
    End Property

    Private Sub SCargueDtbNovedadesNota()
        If IsNothing(MdtbNovedadesNCon) Then
            Dim lstrIdNotaCon = ObjIdNotaConEnt.ToString
            If String.IsNullOrEmpty(lstrIdNotaCon) Then lstrIdNotaCon = "0"
            Dim lstrTabla = ClsNovedad.SstrNombreTabla
            Dim lstrCamposSelect = {"*", "'' AS NroFac", "'' AS Detalle"}
            Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsIdTipoDocOrigenByt.SstrNombreCampoBd &
                    " = " & EnuTipoDocOri.EnuNotaCon & " AND " &
                    ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd & " = '" &
                    ObjPrefijo_NotaConStr.ObjValorPro & "' AND " & ClsIdDocOrigenEnt.SstrNombreCampoBd & " = " &
                    lstrIdNotaCon
            Dim lstrIndice = {{ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd, "ASC"},
                              {ClsIdDocOrigenEnt.SstrNombreCampoBd, "ASC"},
                              {ClsPrefijoFact_NovStr.SstrNombreCampoBd, "ASC"},
                              {ClsIdFactura_NovEnt.SstrNombreCampoBd, "ASC"},
                              {ClsIdNovedadShr.SstrNombreCampoBd, "ASC"}}
            MdtbNovedadesNCon = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamposSelect, lstrIndice, lstrFiltro)
        End If
    End Sub

    Private Sub SComplementeTablaNov()
        Dim lstrPrefFact = String.Empty, lentIdFact = 0
        Dim lstrNroFact = String.Empty, lstrDetalle = String.Empty
        Dim lenuTipoNovedad As EnuTipoNov = EnuTipoNov.None
        Dim ldrwNovedades = MdtbNovedadesNCon.Select
        For Each ldrwNovedad As DataRow In ldrwNovedades
            lstrPrefFact = ClsPanorama.FobjValorCampo(ldrwNovedad(ClsPrefijoFact_NovStr.SstrNombreCampoBd),
                    EnuTipoValor.EnuString)
            lentIdFact = ClsPanorama.FobjValorCampo(ldrwNovedad(ClsIdFactura_NovEnt.SstrNombreCampoBd),
                    EnuTipoValor.EnuInteger)
            lenuTipoNovedad = ClsPanorama.FobjValorCampo(ldrwNovedad(ClsIdTipoNovedadByt.SstrNombreCampoBd),
                    EnuTipoValor.EnuByte)
            lstrNroFact = ClsPanorama.FstrNumeroDcto(lstrPrefFact, lentIdFact)
            Select Case lenuTipoNovedad
                Case EnuTipoNov.EnuCrAnApCap
                    lstrDetalle = My.Resources.NAAAntAplCap
                Case EnuTipoNov.EnuCrAnApInt
                    lstrDetalle = My.Resources.NAAAntAplInt
                Case EnuTipoNov.EnuCrDctoCap
                    lstrDetalle = My.Resources.NAADsctoPP
                Case EnuTipoNov.EnuCrRetFte
                    lstrDetalle = My.Resources.NAAReteFte
                Case EnuTipoNov.EnuCrRetIca
                    lstrDetalle = My.Resources.RCRetIca
                Case EnuTipoNov.EnuCrRetIva
                    lstrDetalle = My.Resources.RCRetIva
                Case EnuTipoNov.EnuRCrAnApCap
                    lstrDetalle = My.Resources.NAARAntAplCap
                Case EnuTipoNov.EnuRCrAnApInt
                    lstrDetalle = My.Resources.NAARAntAplInt
                Case EnuTipoNov.EnuRCrDctoCap
                    lstrDetalle = My.Resources.NAARDsctoPP
                Case EnuTipoNov.EnuRCrRetFte
                    lstrDetalle = My.Resources.NAARReteFte
                Case EnuTipoNov.EnuRCrRetIca
                    lstrDetalle = My.Resources.NAARReteIca
                Case EnuTipoNov.EnuRCrRetIva
                    lstrDetalle = My.Resources.NAARReteIva
            End Select
            ldrwNovedad("NroFac") = lstrNroFact
            ldrwNovedad("Detalle") = lstrDetalle
        Next
    End Sub
#End Region
End Class
#Region "Clases de Propiedad"
Friend Class ClsFechaAnulacion_NotaConDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaAnulacion"
    Private ReadOnly MobjPadre As ClsNotaCon = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "FechaAnulacion_NotaCon"
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
        Dim ldtmFechaMin = GCDTMFECHANULA
        Dim ldtmFechaMax = GCDTMFECHANULA
        Select Case MobjPadre.EnuEstadoActualizacion
            Case EnuEstadoObjetoDef.enuCreando
                '
            Case EnuEstadoObjetoDef.enuModificando
                If Not ClsOrionCop.BlnFacturando Then
                    If MobjPadre.ObjAnuladoBln.ObjValorPro Then
                        ldtmFechaMin = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
                        ldtmFechaMax = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
                    Else
                        ldtmFechaMin = Date.Today
                        ldtmFechaMax = Now
                    End If
                End If
            Case EnuEstadoObjetoDef.enuConsultando
                ldtmFechaMin = GCDTMFECHANULA
                ldtmFechaMax = Now
        End Select
        HblnEsValido = ClsPanorama.FblnEsValidoFecha(HobjValorNew, ldtmFechaMin, ldtmFechaMax,
                BlnEsRequerido)
    End Sub
    Private Sub EPosSetValor() Handles Me.EvnPosSetValor
        If MobjPadre.EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando AndAlso BlnEsValido Then
            If HobjValorPro <> GCDTMFECHANULA Then
                MobjPadre.ObjValor_NotaConDec.SValide()
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
Friend Class ClsFecha_NotaConDtm
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FechaNota"
    Private ReadOnly MobjPadre As ClsNotaCon = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "FechaNotaCon"
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
        Dim ldtmFechaMin = GCDTMFECHANULA
        Dim ldtmFechaMax = GCDTMFECHAMAXI
        If Not GblnActualizandoApp Then
            ldtmFechaMin = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
            ldtmFechaMax = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
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
Friend Class ClsIdAnticipo_NotaConEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdAnticipo"
    Private ReadOnly MobjPadre As ClsNotaCon = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdAnticipo_NotaCon"
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
Friend Class ClsIdCliente_NotaConDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdTerceroCliente"
    Private MstrNombreCliente As String = String.Empty
    Private ReadOnly MobjPadre As ClsNotaCon = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdTerceroCliente_NotaCon"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Protected Overrides Sub SVaciePropiedad()
        MstrNombreCliente = String.Empty
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, GCDBLMINTERC, GCDBLMAXTERC, BlnEsRequerido)
        If HblnEsValido Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                Dim lobjCliente As ClsCliente = MobjPadre.ObjPadre
                HblnEsValido = (HobjValorNew = lobjCliente.ObjIdClienteDbl.ObjValorPro)
            Else
                HblnEsValido = (HobjValorNew = HobjValorOriginal)
            End If
        End If
    End Sub
    Friend ReadOnly Property StrNombreCliente As String
        Get
            If HblnEsValido AndAlso MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                If String.IsNullOrEmpty(MstrNombreCliente) Then
                    Dim lobjCliente As New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
                    Dim lobjvalorLlave() As Object = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorPro}
                    lobjCliente.SAbra(lobjvalorLlave)
                    MstrNombreCliente = lobjCliente.ObjNombreCompletoStr.ObjValorPro
                End If
            End If
            Return MstrNombreCliente
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
Friend Class ClsIdNotaConEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdNotaCon"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "IdNotaContable"
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
Friend Class ClsIdNotaRCrEnt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdNotaRCr"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Id Nota RRC"
        HenuTipoValor = EnuTipoValor.enuInteger
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = False
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, Integer.MaxValue, BlnEsRequerido,
                EnuTipoValor.enuInteger)
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
Friend Class ClsIdPredioAgrupador_NotaConStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "IdPredioAgrupador"
    Private ReadOnly MobjPadre As ClsNotaCon = Nothing
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        MobjPadre = aobjPadre
        HstrNombre = "IdPredioAgrupador NotaCon"
        HenuTipoValor = EnuTipoValor.enuString
        HshrLongitud = 20
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnEsRequerido = True
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then HobjValorNew = String.Empty
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud,
                BlnEsRequerido)
        If HblnEsValido AndAlso HobjValorNew <> "" Then
            If MobjPadre.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuCreando Then
                Dim lobjLlavePrincipal() As Object = {GshrIdCarpeta, GshrIdCentroUtil, HobjValorNew}
                Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
                lobjPredio.SAbra(lobjLlavePrincipal)
                HblnEsValido = lobjPredio.BlnExiste
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
Friend Class ClsPrefijo_NotaConStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PrefijoNotaCon"
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
                HblnEsValido = (HobjValorNew = GobjParametros.FstrPrefijoDoc(EnuTipoDocOri.EnuNotaCon))
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
Friend Class ClsPrefijo_NotaRCrStr
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "PrefijoNotaRCr"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Prefijo Nota RRC"
        HshrLongitud = 5
        HenuTipoValor = EnuTipoValor.enuString
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
    End Sub
    Public Overrides Sub SValide()
        If IsNothing(HobjValorNew) Then
            HobjValorNew = String.Empty
        End If
        HblnEsValido = ClsPanorama.FblnEsValidoString(HobjValorNew, 0, ShrLongitud, BlnEsRequerido)
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
Friend Class ClsValor_NotaConDec
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "Valor"
    Private ReadOnly MobjPadre As ClsNotaCon = Nothing
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
            Return Format(HobjValorPro, "c")
        End If
    End Function
End Class
#End Region